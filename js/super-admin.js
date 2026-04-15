import { auth, db } from './firebase-config.js';
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

let currentUser = null;
let currentUserData = null;
let allUsers = [];
let allSubmissions = [];
let allAudits = [];
let allRoutingRules = [];
let currentTab = 'global';

const superAdminName = document.getElementById('superAdminName');
const pageTitle = document.getElementById('pageTitle');
const notification = document.getElementById('notification');

function showNotification(message, type = 'info') {
    if (!notification) return;
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    setTimeout(() => { notification.style.display = 'none'; }, 3000);
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

function formatDate(ts) {
    if (!ts) return '-';
    try {
        const d = ts.toDate ? ts.toDate() : new Date(ts);
        return d.toLocaleString('en-NG');
    } catch (_) {
        return '-';
    }
}

function roleHome(role) {
    const r = String(role || '').toLowerCase();
    if (r === 'super_admin') return 'super-admin-dashboard.html';
    if (r === 'admin') return 'admin-dashboard.html';
    if (r === 'reviewer' || r === 'viewer') return 'reviewer-dashboard.html';
    if (r === 'rsa') return 'rsa-dashboard.html';
    if (r === 'payment') return 'payment-dashboard.html';
    return 'dashboard.html';
}

function switchTab(tabId) {
    currentTab = tabId;
    document.querySelectorAll('.nav-item').forEach((n) => n.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((t) => t.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');

    const titles = {
        global: 'Global View',
        admins: 'Admin Management',
        'routing-rules': 'Uploader Routing',
        audit: 'Audit',
        security: 'Security',
        'round-robin': 'Round Robin',
        settings: 'Settings',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Super Admin';
    renderCurrentTab();
}

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
    const adminUsers = allUsers.filter((u) => ['admin', 'super_admin'].includes(String(u.role || '').toLowerCase()));
    if (!adminUsers.length) {
        body.innerHTML = '<tr><td colspan="5" class="no-data">No admin users found</td></tr>';
        return;
    }

    body.innerHTML = adminUsers.map((u) => {
        const role = String(u.role || 'admin').toLowerCase();
        const status = String(u.status || 'active').toLowerCase();
        return `
            <tr>
                <td><strong>${escapeHtml(u.fullName || u.email || 'Unknown')}</strong></td>
                <td>${escapeHtml(u.email || '-')}</td>
                <td>
                    <select id="sa-role-${u.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                        <option value="super_admin" ${role === 'super_admin' ? 'selected' : ''}>Super Admin</option>
                        <option value="admin" ${role === 'admin' ? 'selected' : ''}>Admin</option>
                    </select>
                </td>
                <td>
                    <select id="sa-status-${u.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                        <option value="active" ${status === 'active' ? 'selected' : ''}>Active</option>
                        <option value="deactivated" ${status === 'deactivated' ? 'selected' : ''}>Deactivated</option>
                    </select>
                </td>
                <td>
                    <button class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.saveAdminManager('${u.id}')">
                        <i class="fas fa-save"></i> Save
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function getUsersByRole(role) {
    const r = String(role || '').toLowerCase();
    return allUsers.filter((u) => String(u.role || '').toLowerCase() === r && String(u.status || 'active').toLowerCase() !== 'deactivated');
}

function getActiveUsersByRoles(roles = []) {
    const set = new Set(roles.map((r) => String(r || '').toLowerCase()));
    return allUsers.filter((u) => {
        const role = String(u.role || '').toLowerCase();
        const status = String(u.status || 'active').toLowerCase();
        return set.has(role) && status !== 'deactivated';
    });
}

function findRoutingRuleForUploader(uploaderEmail) {
    const normalized = normalizeEmail(uploaderEmail);
    return allRoutingRules.find((r) => normalizeEmail(r.uploaderEmail) === normalized) || null;
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
    const body = document.getElementById('routingRulesTableBody');
    if (!body) return;

    const uploaderUsers = getUsersByRole('uploader');
    const reviewerUsers = getActiveUsersByRoles(['reviewer', 'viewer']);
    const rsaUsers = getUsersByRole('rsa');
    const paymentUsers = getUsersByRole('payment');

    if (!uploaderUsers.length) {
        body.innerHTML = '<tr><td colspan="7" class="no-data">No uploader users found</td></tr>';
        return;
    }

    body.innerHTML = uploaderUsers.map((uploader) => {
        const uploaderEmail = normalizeEmail(uploader.email);
        const rule = findRoutingRuleForUploader(uploaderEmail) || {};
        const selectedReviewer = normalizeEmail(rule.reviewerEmail);
        const selectedRsa = normalizeEmail(rule.rsaEmail);
        const selectedPayment = normalizeEmail(rule.paymentEmail);
        const updatedAt = formatDate(rule.updatedAt || rule.createdAt);

        return `
            <tr>
                <td><strong>${escapeHtml(uploader.fullName || uploaderEmail || 'Unknown')}</strong></td>
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
                    <button class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.saveUploaderRoutingRule('${uploader.id}')">
                        <i class="fas fa-save"></i> Save
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function renderAudit() {
    const body = document.getElementById('superAuditTableBody');
    if (!body) return;
    if (!allAudits.length) {
        body.innerHTML = '<tr><td colspan="4" class="no-data">No audit records</td></tr>';
        return;
    }
    body.innerHTML = allAudits.slice(0, 200).map((a) => `
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
    const portalUrl = document.getElementById('settingPortalUrl');
    const emailProvider = document.getElementById('settingEmailProvider');
    const maintenance = document.getElementById('settingMaintenanceMode');
    if (portalUrl) portalUrl.value = data.portalUrl || window.location.origin;
    if (emailProvider) emailProvider.value = data.emailProvider || 'emailjs';
    if (maintenance) maintenance.value = data.maintenanceMode ? 'true' : 'false';
}

function renderCurrentTab() {
    if (currentTab === 'global') return renderGlobalView();
    if (currentTab === 'admins') return renderAdminManagement();
    if (currentTab === 'routing-rules') return renderRoutingRules();
    if (currentTab === 'audit') return renderAudit();
    if (currentTab === 'security') return renderSecurity();
    if (currentTab === 'round-robin') return renderRoundRobin();
    if (currentTab === 'settings') return loadSettings();
}

window.saveUploaderRoutingRule = async (uploaderUserId) => {
    const uploader = allUsers.find((u) => u.id === uploaderUserId);
    if (!uploader) return showNotification('Uploader not found', 'error');

    const uploaderEmail = normalizeEmail(uploader.email);
    if (!uploaderEmail) return showNotification('Uploader email is missing', 'error');

    const reviewerEmail = normalizeEmail(document.getElementById(`route-reviewer-${uploaderUserId}`)?.value);
    const rsaEmail = normalizeEmail(document.getElementById(`route-rsa-${uploaderUserId}`)?.value);
    const paymentEmail = normalizeEmail(document.getElementById(`route-payment-${uploaderUserId}`)?.value);
    const enabled = Boolean(reviewerEmail || rsaEmail || paymentEmail);
    const existingRule = findRoutingRuleForUploader(uploaderEmail);

    try {
        const payload = {
            uploaderEmail,
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
            reviewerEmail: reviewerEmail || '',
            rsaEmail: rsaEmail || '',
            paymentEmail: paymentEmail || '',
            enabled,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification('Uploader routing rule saved', 'success');
    } catch (error) {
        showNotification('Failed to save routing rule', 'error');
    }
};

window.saveAdminManager = async (userId) => {
    const roleEl = document.getElementById(`sa-role-${userId}`);
    const statusEl = document.getElementById(`sa-status-${userId}`);
    const role = String(roleEl?.value || '').trim();
    const status = String(statusEl?.value || '').trim();
    if (!role) return showNotification('Role is required', 'error');

    try {
        await updateDoc(doc(db, 'users', userId), {
            role,
            status,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        });
        await addDoc(collection(db, 'audit'), {
            action: 'admin_management_updated',
            userId,
            newRole: role,
            newStatus: status,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('Admin updated successfully', 'success');
    } catch (error) {
        showNotification('Failed to update admin user', 'error');
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

window.saveSuperSettings = async () => {
    const portalUrl = String(document.getElementById('settingPortalUrl')?.value || '').trim();
    const emailProvider = String(document.getElementById('settingEmailProvider')?.value || '').trim();
    const maintenanceMode = String(document.getElementById('settingMaintenanceMode')?.value || 'false') === 'true';
    try {
        await setDoc(doc(db, 'settings', 'system'), {
            portalUrl,
            emailProvider,
            maintenanceMode,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        }, { merge: true });
        await addDoc(collection(db, 'audit'), {
            action: 'super_admin_settings_updated',
            portalUrl,
            emailProvider,
            maintenanceMode,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('Settings saved', 'success');
    } catch (error) {
        showNotification('Failed to save settings', 'error');
    }
};

window.signOutUser = async () => {
    try { await signOut(auth); } catch (_) {}
    window.location.href = 'index.html';
};

function setupRealtimeData() {
    onSnapshot(query(collection(db, 'users')), (snap) => {
        allUsers = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        renderCurrentTab();
        if (currentTab !== 'global') renderGlobalView();
    });

    onSnapshot(query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc')), (snap) => {
        allSubmissions = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        renderCurrentTab();
        if (currentTab !== 'global') renderGlobalView();
    });

    onSnapshot(query(collection(db, 'audit'), orderBy('timestamp', 'desc'), limit(200)), (snap) => {
        allAudits = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        if (currentTab === 'audit') renderAudit();
    });

    onSnapshot(query(collection(db, 'uploaderRoutingRules')), (snap) => {
        allRoutingRules = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        if (currentTab === 'routing-rules') renderRoutingRules();
    });
}

document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        switchTab(item.dataset.tab);
    });
});

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
        const byUid = query(collection(db, 'users'), where('uid', '==', user.uid));
        let snap = await getDocs(byUid);
        if (snap.empty && user.email) {
            const byEmail = query(collection(db, 'users'), where('email', '==', user.email.toLowerCase()));
            snap = await getDocs(byEmail);
        }
        if (snap.empty) return void (window.location.href = 'index.html');

        currentUserData = snap.docs[0].data() || {};
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
