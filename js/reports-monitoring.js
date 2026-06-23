import { auth, db } from './firebase-config.js';
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    doc,
    onSnapshot,
    serverTimestamp,
    updateDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { formatAppDateTime } from './shared/app-time.js';
import { getCurrentUserProfile as getCurrentUserProfileShared } from './shared/user-directory.js?v=20260518a';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionCurrentStageEntryAt
} from './shared/submission-stage.js?v=20260609a';

let currentUser = null;
let currentUserData = null;
let allUsers = [];
let allSubmissions = [];
let currentTab = 'overview';
let usersListenerStarted = false;
let submissionsListenerStarted = false;

const pageTitle = document.getElementById('pageTitle');
const monitorName = document.getElementById('monitorName');
const notification = document.getElementById('notification');
const usersSearch = document.getElementById('usersSearch');
const usersRoleFilter = document.getElementById('usersRoleFilter');
const usersStatusFilter = document.getElementById('usersStatusFilter');
const applicationsSearch = document.getElementById('applicationsSearch');
const applicationsStatusFilter = document.getElementById('applicationsStatusFilter');
const applicationsStageFilter = document.getElementById('applicationsStageFilter');
const applicationDetailsModal = document.getElementById('applicationDetailsModal');

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

function showNotification(message, type = 'info') {
    if (!notification) return;
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    clearTimeout(showNotification._timer);
    showNotification._timer = setTimeout(() => {
        notification.style.display = 'none';
    }, 3000);
}

function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
}

function formatDate(value) {
    return formatAppDateTime(value, '-');
}

function roleLabel(role) {
    const normalized = String(role || '').trim().toLowerCase();
    if (normalized === 'super_admin') return 'Super Admin';
    if (normalized === 'admin') return 'Admin';
    if (normalized === 'reports_monitoring') return 'Reports Monitoring';
    if (normalized === 'reviewer') return 'Reviewer';
    if (normalized === 'rsa') return 'RSA';
    if (normalized === 'payment') return 'Payment';
    return 'Uploader';
}

function statusLabel(status) {
    const normalized = String(status || '').trim().toLowerCase();
    if (!normalized) return 'Unknown';
    return normalized.replaceAll('_', ' ').replace(/\b\w/g, (char) => char.toUpperCase());
}

function getApplicationStage(sub = {}) {
    const status = String(sub.status || '').trim().toLowerCase();
    if (sub.isDraft || status === 'draft') return 'Draft';
    if (['paid', 'cleared', 'sent_to_pfa', 'rsa_submitted'].includes(status) || sub.finalSubmitted === true || sub.rsaSubmitted === true) return 'Payment';
    if (['processing_to_pfa', 'approved', 'rejected_by_rsa'].includes(status) || normalizeEmail(sub.assignedToRSA)) return 'RSA';
    if (['pending', 'rejected'].includes(status) || normalizeEmail(sub.assignedTo) || normalizeEmail(sub.reviewedBy)) return 'Review';
    return 'Closed';
}

function getWorkflowSnapshotItems() {
    return [
        {
            label: 'Pending Review',
            count: allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length,
            note: 'Applications waiting in reviewer flow.'
        },
        {
            label: 'RSA Processing',
            count: allSubmissions.filter((sub) => ['approved', 'processing_to_pfa', 'rejected_by_rsa'].includes(String(sub.status || '').toLowerCase())).length,
            note: 'Applications currently in or returned from RSA stage.'
        },
        {
            label: 'Sent to PFA',
            count: allSubmissions.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase()) || sub.finalSubmitted === true || sub.rsaSubmitted === true).length,
            note: 'Applications already forwarded for payment processing.'
        },
        {
            label: 'Paid',
            count: allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'paid').length,
            note: 'Applications with payment confirmed.'
        },
        {
            label: 'Cleared',
            count: allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'cleared').length,
            note: 'Applications fully settled and closed financially.'
        }
    ];
}

function getOfficerWorkloadRows() {
    const officerRoles = new Set(['reviewer', 'rsa', 'payment']);
    return allUsers
        .filter((user) => officerRoles.has(String(user.role || '').toLowerCase()))
        .map((user) => {
            const email = normalizeEmail(user.email);
            const role = String(user.role || '').toLowerCase();
            let assignedCount = 0;
            let completedCount = 0;

            if (role === 'reviewer') {
                assignedCount = allSubmissions.filter((sub) => normalizeEmail(sub.assignedTo) === email).length;
                completedCount = allSubmissions.filter((sub) => normalizeEmail(sub.reviewedBy) === email).length;
            } else if (role === 'rsa') {
                assignedCount = allSubmissions.filter((sub) => normalizeEmail(sub.assignedToRSA) === email).length;
                completedCount = allSubmissions.filter((sub) => normalizeEmail(sub.assignedToRSA) === email && ['processing_to_pfa', 'sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase())).length;
            } else if (role === 'payment') {
                assignedCount = allSubmissions.filter((sub) => normalizeEmail(sub.assignedToPayment) === email).length;
                completedCount = allSubmissions.filter((sub) => normalizeEmail(sub.assignedToPayment) === email && ['paid', 'cleared'].includes(String(sub.status || '').toLowerCase())).length;
            }

            return {
                fullName: user.fullName || user.email || 'Unknown',
                role: roleLabel(role),
                assignedCount,
                completedCount,
                status: statusLabel(user.status || 'active')
            };
        })
        .sort((a, b) => b.assignedCount - a.assignedCount || a.fullName.localeCompare(b.fullName));
}

function getExceptionRows() {
    return allSubmissions
        .map((sub) => {
            const status = String(sub.status || '').toLowerCase();
            if (status === 'pending' && !normalizeEmail(sub.assignedTo)) {
                return { sub, issue: 'No reviewer assigned' };
            }
            if (['approved', 'processing_to_pfa'].includes(status) && !normalizeEmail(sub.assignedToRSA)) {
                return { sub, issue: 'No RSA officer assigned' };
            }
            if ((['sent_to_pfa', 'rsa_submitted'].includes(status) || sub.finalSubmitted === true || sub.rsaSubmitted === true) && !normalizeEmail(sub.assignedToPayment)) {
                return { sub, issue: 'No payment officer assigned' };
            }
            if (!String(sub.customerName || '').trim()) {
                return { sub, issue: 'Customer name missing' };
            }
            if (!String(sub.agentName || '').trim()) {
                return { sub, issue: 'Agent not linked' };
            }
            return null;
        })
        .filter(Boolean)
        .slice(0, 100);
}

function filteredUsers() {
    const search = String(usersSearch?.value || '').trim().toLowerCase();
    const role = String(usersRoleFilter?.value || '').trim().toLowerCase();
    const status = String(usersStatusFilter?.value || '').trim().toLowerCase();

    return allUsers.filter((user) => {
        const matchesSearch = !search || [
            user.fullName,
            user.email,
            user.location,
            user.whatsappNumber,
            user.phone,
            user.role
        ].some((value) => String(value || '').toLowerCase().includes(search));
        const matchesRole = !role || String(user.role || '').toLowerCase() === role;
        const matchesStatus = !status || String(user.status || '').toLowerCase() === status;
        return matchesSearch && matchesRole && matchesStatus;
    });
}

function filteredApplications() {
    const search = String(applicationsSearch?.value || '').trim().toLowerCase();
    const status = String(applicationsStatusFilter?.value || '').trim().toLowerCase();
    const stage = String(applicationsStageFilter?.value || '').trim();

    return allSubmissions.filter((sub) => {
        const submissionStage = getApplicationStage(sub);
        const matchesSearch = !search || [
            sub.customerName,
            sub.uploadedBy,
            sub.agentName,
            sub.id,
            sub.assignedTo,
            sub.assignedToRSA,
            sub.assignedToPayment
        ].some((value) => String(value || '').toLowerCase().includes(search));
        const matchesStatus = !status || String(sub.status || '').toLowerCase() === status;
        const matchesStage = !stage || submissionStage === stage;
        return matchesSearch && matchesStatus && matchesStage;
    });
}

function setCountBadge(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = String(value);
}

function renderOverview() {
    const sentCount = allSubmissions.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase()) || sub.finalSubmitted === true || sub.rsaSubmitted === true).length;
    const paidClearedCount = allSubmissions.filter((sub) => ['paid', 'cleared'].includes(String(sub.status || '').toLowerCase())).length;

    setCountBadge('overviewUsersCount', allUsers.length);
    setCountBadge('overviewApplicationsCount', allSubmissions.length);
    setCountBadge('overviewSentCount', sentCount);
    setCountBadge('overviewPaidClearedCount', paidClearedCount);
    setCountBadge('usersCountBadge', allUsers.length);
    setCountBadge('applicationsCountBadge', allSubmissions.length);

    const workflowBody = document.getElementById('overviewWorkflowBody');
    if (workflowBody) {
        const rows = getWorkflowSnapshotItems();
        workflowBody.innerHTML = rows.map((item) => `
            <tr>
                <td>${escapeHtml(item.label)}</td>
                <td><strong>${item.count}</strong></td>
                <td>${escapeHtml(item.note)}</td>
            </tr>
        `).join('');
    }

    const recentBody = document.getElementById('recentApplicationsBody');
    if (recentBody) {
        const rows = allSubmissions
            .slice()
            .sort((a, b) => {
                const aMs = getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a));
                const bMs = getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b));
                return bMs - aMs;
            })
            .slice(0, 10);

        recentBody.innerHTML = rows.length ? rows.map((sub) => `
            <tr>
                <td>${escapeHtml(sub.customerName || 'Unknown')}</td>
                <td>${escapeHtml(statusLabel(sub.status || '-'))}</td>
                <td>${escapeHtml(getApplicationStage(sub))}</td>
                <td>${escapeHtml(sub.uploadedBy || '-')}</td>
                <td>${escapeHtml([sub.assignedTo || '-', sub.assignedToRSA || '-', sub.assignedToPayment || '-'].join(' / '))}</td>
                <td>${escapeHtml(formatDate(getSubmissionCurrentStageEntryAt(sub)))}</td>
                <td><button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button></td>
            </tr>
        `).join('') : '<tr><td colspan="7" class="no-data">No applications available</td></tr>';
    }
}

function renderUsers() {
    const body = document.getElementById('usersTableBody');
    if (!body) return;
    const rows = filteredUsers()
        .slice()
        .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));

    body.innerHTML = rows.length ? rows.map((user) => `
        <tr>
            <td><strong>${escapeHtml(user.fullName || user.email || 'Unknown')}</strong></td>
            <td>${escapeHtml(user.email || '-')}</td>
            <td>${escapeHtml(roleLabel(user.role))}</td>
            <td>${escapeHtml(statusLabel(user.status || 'active'))}</td>
            <td>${escapeHtml(user.location || '-')}</td>
            <td>${escapeHtml(user.whatsappNumber || user.phone || '-')}</td>
            <td>${escapeHtml(formatDate(user.createdAt))}</td>
        </tr>
    `).join('') : '<tr><td colspan="7" class="no-data">No users found</td></tr>';
}

function renderApplications() {
    const body = document.getElementById('applicationsTableBody');
    if (!body) return;
    const rows = filteredApplications()
        .slice()
        .sort((a, b) => {
            const aMs = getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a));
            const bMs = getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b));
            return bMs - aMs;
        });

    body.innerHTML = rows.length ? rows.map((sub) => `
        <tr>
            <td>${escapeHtml(sub.customerName || 'Unknown')}</td>
            <td>${escapeHtml(sub.id || '-')}</td>
            <td>${escapeHtml(statusLabel(sub.status || '-'))}</td>
            <td>${escapeHtml(getApplicationStage(sub))}</td>
            <td>${escapeHtml(sub.uploadedBy || '-')}</td>
            <td>${escapeHtml(sub.assignedTo || sub.reviewedBy || '-')}</td>
            <td>${escapeHtml(sub.assignedToRSA || '-')}</td>
            <td>${escapeHtml(sub.assignedToPayment || '-')}</td>
            <td>${escapeHtml(formatDate(getSubmissionCurrentStageEntryAt(sub)))}</td>
            <td><button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button></td>
        </tr>
    `).join('') : '<tr><td colspan="10" class="no-data">No applications found</td></tr>';
}

function renderReports() {
    const rsaReportsBody = document.getElementById('rsaReportsBody');
    if (rsaReportsBody) {
        const rows = [
            {
                label: 'Processing to PFA Report',
                count: allSubmissions.filter((sub) => ['approved', 'processing_to_pfa'].includes(String(sub.status || '').toLowerCase())).length,
                description: 'Applications currently being handled in RSA before final submission.'
            },
            {
                label: 'Final Submitted Report',
                count: allSubmissions.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase()) || sub.finalSubmitted === true || sub.rsaSubmitted === true).length,
                description: 'Applications already forwarded to payment workflow.'
            },
            {
                label: 'Rejected by RSA Report',
                count: allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'rejected_by_rsa').length,
                description: 'Applications returned by RSA for correction.'
            }
        ];
        rsaReportsBody.innerHTML = rows.map((item) => `
            <tr>
                <td>${escapeHtml(item.label)}</td>
                <td><strong>${item.count}</strong></td>
                <td>${escapeHtml(item.description)}</td>
            </tr>
        `).join('');
    }

    const paymentReportsBody = document.getElementById('paymentReportsBody');
    if (paymentReportsBody) {
        const rows = [
            {
                label: 'Sent to PFA Queue Report',
                count: allSubmissions.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase()) || sub.finalSubmitted === true || sub.rsaSubmitted === true).length,
                description: 'Applications waiting for payment processing.'
            },
            {
                label: 'Paid Report',
                count: allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'paid').length,
                description: 'Applications marked paid by payment officers.'
            },
            {
                label: 'Cleared Report',
                count: allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'cleared').length,
                description: 'Applications fully settled and cleared.'
            }
        ];
        paymentReportsBody.innerHTML = rows.map((item) => `
            <tr>
                <td>${escapeHtml(item.label)}</td>
                <td><strong>${item.count}</strong></td>
                <td>${escapeHtml(item.description)}</td>
            </tr>
        `).join('');
    }

    const workloadBody = document.getElementById('officerWorkloadBody');
    if (workloadBody) {
        const rows = getOfficerWorkloadRows();
        workloadBody.innerHTML = rows.length ? rows.map((row) => `
            <tr>
                <td>${escapeHtml(row.fullName)}</td>
                <td>${escapeHtml(row.role)}</td>
                <td><strong>${row.assignedCount}</strong></td>
                <td>${row.completedCount}</td>
                <td>${escapeHtml(row.status)}</td>
            </tr>
        `).join('') : '<tr><td colspan="5" class="no-data">No officer workload data available</td></tr>';
    }

    const exceptionsBody = document.getElementById('exceptionsTableBody');
    if (exceptionsBody) {
        const rows = getExceptionRows();
        exceptionsBody.innerHTML = rows.length ? rows.map(({ sub, issue }) => `
            <tr>
                <td>${escapeHtml(sub.customerName || 'Unknown')}</td>
                <td>${escapeHtml(issue)}</td>
                <td>${escapeHtml(statusLabel(sub.status || '-'))}</td>
                <td>${escapeHtml(getApplicationStage(sub))}</td>
                <td><button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button></td>
            </tr>
        `).join('') : '<tr><td colspan="5" class="no-data">No workflow exceptions detected</td></tr>';
    }
}

function renderProfile() {
    const fullName = currentUserData?.fullName || currentUser?.displayName || currentUser?.email || 'N/A';
    const email = currentUserData?.email || currentUser?.email || 'N/A';
    const whatsapp = currentUserData?.whatsappNumber || currentUserData?.phone || '-';
    const location = currentUserData?.location || '-';
    const role = roleLabel(currentUserData?.role || 'reports_monitoring');
    const status = statusLabel(currentUserData?.status || 'active');
    const registeredAt = formatDate(currentUserData?.createdAt);

    if (monitorName) monitorName.textContent = fullName;
    const mappings = {
        profileName: fullName,
        profileEmail: email,
        profileWhatsapp: whatsapp,
        profileLocation: location,
        profileRole: role,
        profileStatus: status,
        profileRegisteredAt: registeredAt
    };

    Object.entries(mappings).forEach(([id, value]) => {
        const el = document.getElementById(id);
        if (el) el.textContent = value;
    });
}

function renderCurrentTab() {
    renderOverview();
    renderUsers();
    renderApplications();
    renderReports();
}

function switchTab(tabId) {
    currentTab = tabId;
    ensureDataForTab(tabId);
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');

    const titles = {
        overview: 'Overview',
        users: 'Users',
        applications: 'Applications',
        reports: 'Reports',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Reports & Monitoring';
}

function ensureDataForTab(tabId) {
    if (tabId === 'overview' || tabId === 'users' || tabId === 'reports') loadUsers();
    if (tabId === 'overview' || tabId === 'applications' || tabId === 'reports') loadSubmissions();
}

function toggleSidebar(open) {
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('sidebarOverlay');
    sidebar?.classList.toggle('active', open);
    overlay?.classList.toggle('active', open);
    const menuToggle = document.getElementById('menuToggle');
    if (menuToggle) menuToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
}

function openApplicationDetailsModal(submissionId) {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const rows = [
        ['Application ID', sub.id || '-'],
        ['Customer Name', sub.customerName || '-'],
        ['Status', statusLabel(sub.status || '-')],
        ['Stage', getApplicationStage(sub)],
        ['Uploader', sub.uploadedBy || '-'],
        ['Assigned Reviewer', sub.assignedTo || sub.reviewedBy || '-'],
        ['Assigned RSA', sub.assignedToRSA || '-'],
        ['Assigned Payment', sub.assignedToPayment || '-'],
        ['Agent Name', sub.agentName || '-'],
        ['PFA', sub?.customerDetails?.pfa || sub.pfa || '-'],
        ['Uploaded At', formatDate(sub.uploadedAt)],
        ['Updated At', formatDate(sub.updatedAt || sub.uploadedAt)]
    ];

    const title = document.getElementById('applicationDetailsTitle');
    const body = document.getElementById('applicationDetailsBody');
    if (title) title.textContent = `Application Details - ${sub.customerName || 'Customer'}`;
    if (body) {
        body.innerHTML = rows.map(([label, value]) => `
            <tr>
                <th style="width:240px;background:#f8fafc;">${escapeHtml(label)}</th>
                <td>${escapeHtml(value)}</td>
            </tr>
        `).join('');
    }
    applicationDetailsModal?.classList.add('active');
}

function closeApplicationDetailsModal() {
    applicationDetailsModal?.classList.remove('active');
}

window.openMonitoringApplicationDetails = openApplicationDetailsModal;
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
        window.location.href = 'index.html';
    } catch (_) {
        showNotification('Failed to sign out', 'error');
    }
};

function bindEvents() {
    document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            switchTab(item.dataset.tab || 'overview');
        });
    });

    [usersSearch, usersRoleFilter, usersStatusFilter, applicationsSearch, applicationsStatusFilter, applicationsStageFilter]
        .forEach((el) => el?.addEventListener('input', renderCurrentTab));
    [usersRoleFilter, usersStatusFilter, applicationsStatusFilter, applicationsStageFilter]
        .forEach((el) => el?.addEventListener('change', renderCurrentTab));

    document.getElementById('signOutBtnSidebar')?.addEventListener('click', window.signOutUser);
    document.getElementById('signOutBtnMobile')?.addEventListener('click', window.signOutUser);
    document.getElementById('forceRefreshBtn')?.addEventListener('click', () => window.location.reload());
    document.getElementById('forceRefreshBtnMobile')?.addEventListener('click', () => window.location.reload());

    document.getElementById('menuToggle')?.addEventListener('click', () => toggleSidebar(true));
    document.getElementById('sidebarClose')?.addEventListener('click', () => toggleSidebar(false));
    document.getElementById('sidebarOverlay')?.addEventListener('click', () => toggleSidebar(false));

    document.getElementById('closeApplicationDetailsModalBtn')?.addEventListener('click', closeApplicationDetailsModal);
    document.getElementById('closeApplicationDetailsModalFooterBtn')?.addEventListener('click', closeApplicationDetailsModal);
    window.addEventListener('click', (e) => {
        if (e.target === applicationDetailsModal) closeApplicationDetailsModal();
    });
}

function loadUsers() {
    if (usersListenerStarted) return;
    usersListenerStarted = true;
    onSnapshot(collection(db, 'users'), (snapshot) => {
        allUsers = snapshot.docs
            .map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }))
            .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));
        renderCurrentTab();
    }, () => {
        showNotification('Failed to load users', 'error');
    });
}

function loadSubmissions() {
    if (submissionsListenerStarted) return;
    submissionsListenerStarted = true;
    onSnapshot(collection(db, 'submissions'), (snapshot) => {
        allSubmissions = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        renderCurrentTab();
    }, () => {
        showNotification('Failed to load applications', 'error');
    });
}

auth.onAuthStateChanged(async (user) => {
    if (!user) {
        window.location.href = 'index.html';
        return;
    }

    currentUser = user;

    try {
        const userData = await getCurrentUserProfileShared(db, user);
        if (!userData) {
            showNotification('User profile not found', 'error');
            window.location.href = 'index.html';
            return;
        }

        const role = String(userData.role || '').toLowerCase();
        if (role !== 'reports_monitoring') {
            if (role === 'super_admin') window.location.href = 'super-admin-dashboard.html';
            else if (role === 'admin') window.location.href = 'admin-dashboard.html';
            else if (role === 'reviewer') window.location.href = 'reviewer-dashboard.html';
            else if (role === 'rsa') window.location.href = 'rsa-dashboard.html';
            else if (role === 'payment') window.location.href = 'payment-dashboard.html';
            else window.location.href = 'dashboard.html';
            return;
        }

        currentUserData = userData;
        renderProfile();
        bindEvents();
        switchTab('overview');
    } catch (_) {
        showNotification('Could not validate session', 'error');
        window.location.href = 'index.html';
    }
});
