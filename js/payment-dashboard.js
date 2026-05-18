import { auth, db } from './firebase-config.js';
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    addDoc,
    query,
    where,
    getDocs,
    onSnapshot,
    updateDoc,
    doc,
    orderBy,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { notifyStatusChangePush } from './status-push.js';
import { formatAppDateTime } from './shared/app-time.js';
import {
    getSubmissionCommissionAmount,
    resolveSubmissionCommissionRate
} from './shared/commission-config.js?v=20260507a';
import { getCurrentUserProfile as getCurrentUserProfileShared } from './shared/user-directory.js?v=20260518a';

let currentUser = null;
let currentUserData = null;
let allSubmissions = [];
let paymentLeaveHistoryLoaded = false;
let paymentMyLeaveHistory = [];
let paymentReliefLeaveHistory = [];
const uploaderNameCache = new Map();

const pageTitle = document.getElementById('pageTitle');
const paymentUserName = document.getElementById('paymentUserName');
const paymentPendingCount = document.getElementById('paymentPendingCount');
const paidCustomerCount = document.getElementById('paidCustomerCount');
const clearedCustomerCount = document.getElementById('clearedCustomerCount');
const paymentsTableBody = document.getElementById('paymentsTableBody');
const paidCustomersTableBody = document.getElementById('paidCustomersTableBody');
const clearedCustomersTableBody = document.getElementById('clearedCustomersTableBody');
const dashboardSentToPfaCount = document.getElementById('dashboardSentToPfaCount');
const dashboardPaidCount = document.getElementById('dashboardPaidCount');
const dashboardClearedCount = document.getElementById('dashboardClearedCount');
const dashboardTotalCommission = document.getElementById('dashboardTotalCommission');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
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

function formatDateValue(value) {
    return formatAppDateTime(value, '-');
}

function getSubmissionFinancials(sub) {
    const details = sub?.customerDetails || {};
    const rsaBalance = parseMoney(details.rsaBalance || sub?.rsaBalance || 0);
    const computed25 = roundDownToNearestThousand(rsaBalance * 0.25);
    const stored25 = parseMoney(details.rsa25Percent || sub?.rsa25Percent || 0);
    const twentyFive = stored25 ? roundDownToNearestThousand(stored25) : computed25;
    const commissionRate = resolveSubmissionCommissionRate(sub);
    const commission2 = hasCommissionEligibleAgent(sub) ? getSubmissionCommissionAmount(sub, twentyFive) : 0;
    const pfa = String(details.pfa || sub?.pfa || '').trim() || '-';
    return { pfa, twentyFive, commission2, commissionRate };
}

function hasCommissionEligibleAgent(sub) {
    return Boolean(String(sub?.agentId || '').trim() || String(sub?.agentName || '').trim());
}

function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
}

function getUploaderDisplayName(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return '-';
    return uploaderNameCache.get(normalized) || normalized.split('@')[0] || normalized;
}

async function primeUploaderNames(records = []) {
    const emails = [...new Set(records.map((sub) => normalizeEmail(sub?.uploadedBy)).filter(Boolean))];
    const missingEmails = emails.filter((email) => !uploaderNameCache.has(email));
    if (!missingEmails.length) return;

    try {
        const usersSnap = await getDocs(collection(db, 'users'));
        usersSnap.docs.forEach((docSnap) => {
            const data = docSnap.data() || {};
            const email = normalizeEmail(data.email);
            if (!email) return;
            uploaderNameCache.set(email, String(data.fullName || email.split('@')[0] || email).trim());
        });
    } catch (_) {
        missingEmails.forEach((email) => {
            if (!uploaderNameCache.has(email)) {
                uploaderNameCache.set(email, email.split('@')[0] || email);
            }
        });
    }
}

function toSafeDomId(value) {
    return String(value || '').replace(/[^a-zA-Z0-9_-]+/g, '_');
}

function getAgentPaymentKey(sub) {
    const agentId = String(sub?.agentId || '').trim();
    if (agentId) return `agent:${agentId}`;
    const agentName = String(sub?.agentName || '').trim().toLowerCase();
    const uploaderEmail = normalizeEmail(sub?.uploadedBy);
    return `fallback:${uploaderEmail}::${agentName || 'no-agent'}`;
}

function buildAgentPaymentGroups(records = []) {
    const groups = new Map();
    records.forEach((sub) => {
        const key = getAgentPaymentKey(sub);
        const existing = groups.get(key) || {
            key,
            agentId: String(sub?.agentId || '').trim(),
            agentName: String(sub?.agentName || '').trim() || 'No Agent',
            uploaderEmail: normalizeEmail(sub?.uploadedBy),
            uploaderName: getUploaderDisplayName(sub?.uploadedBy),
            agentAccountNumber: String(sub?.agentAccountNumber || '').trim() || '-',
            agentAccountBank: String(sub?.agentAccountBank || '').trim() || '-',
            submissions: [],
            customerNames: new Set(),
            pfas: new Set(),
            total25: 0,
            totalCommission: 0,
            latestPaidAt: null,
            latestClearedAt: null,
            latestQueueAt: null,
            hasCommissionEligibleAgent: false
        };
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        existing.submissions.push(sub);
        existing.hasCommissionEligibleAgent = existing.hasCommissionEligibleAgent || hasCommissionEligibleAgent(sub);
        existing.customerNames.add(String(sub?.customerName || 'Unknown'));
        if (pfa && pfa !== '-') existing.pfas.add(pfa);
        existing.total25 += twentyFive;
        existing.totalCommission += commission2;

        const paidAtMs = sub?.paidAt?.toMillis ? sub.paidAt.toMillis() : new Date(sub?.paidAt || 0).getTime();
        const clearedAtMs = sub?.clearedAt?.toMillis ? sub.clearedAt.toMillis() : new Date(sub?.clearedAt || 0).getTime();
        const queueAtMs = sub?.rsaSubmittedAt?.toMillis ? sub.rsaSubmittedAt.toMillis() : new Date(sub?.rsaSubmittedAt || sub?.updatedAt || 0).getTime();
        if (Number.isFinite(paidAtMs) && paidAtMs > 0 && (!existing.latestPaidAt || paidAtMs > existing.latestPaidAt)) existing.latestPaidAt = paidAtMs;
        if (Number.isFinite(clearedAtMs) && clearedAtMs > 0 && (!existing.latestClearedAt || clearedAtMs > existing.latestClearedAt)) existing.latestClearedAt = clearedAtMs;
        if (Number.isFinite(queueAtMs) && queueAtMs > 0 && (!existing.latestQueueAt || queueAtMs > existing.latestQueueAt)) existing.latestQueueAt = queueAtMs;

        groups.set(key, existing);
    });

    return Array.from(groups.values())
        .map((group) => ({
            ...group,
            customerCount: group.submissions.length,
            customerNames: Array.from(group.customerNames),
            pfas: Array.from(group.pfas)
        }))
        .sort((a, b) => (b.latestQueueAt || b.latestPaidAt || b.latestClearedAt || 0) - (a.latestQueueAt || a.latestPaidAt || a.latestClearedAt || 0));
}

function renderAgentBreakdownTable(group, mode = 'queue') {
    const rows = group.submissions.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const status = String(sub?.status || '').toLowerCase();
        const statusLabel = status === 'cleared' ? 'Cleared' : status === 'paid' ? 'Paid' : 'Sent to PFA';
        const dateLabel = mode === 'cleared'
            ? formatDateValue(sub?.clearedAt || sub?.updatedAt)
            : mode === 'paid'
                ? formatDateValue(sub?.paidAt || sub?.updatedAt)
                : formatDateValue(sub?.rsaSubmittedAt || sub?.updatedAt);

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(sub.agentName || group.agentName || 'No Agent')}</td>
                <td>${escapeHtml(pfa)}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td><span class="status-badge status-approved">${escapeHtml(statusLabel)}</span></td>
                <td>${escapeHtml(dateLabel)}</td>
                <td><button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');

    return `
        <div class="agent-breakdown-panel">
            <div class="agent-breakdown-meta">
                <strong>${escapeHtml(group.agentName)}</strong>
                <span>${group.customerCount} customer(s)</span>
                <span>${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</span>
            </div>
            <div class="table-container">
                <table class="customers-table agent-breakdown-table">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Agent</th>
                            <th>PFA</th>
                            <th>25% Balance</th>
                            <th>Commission Amount</th>
                            <th>Status</th>
                            <th>Date</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
        </div>
    `;
}

function renderProfile() {
    const fullName = currentUserData?.fullName || currentUser?.displayName || currentUser?.email || 'N/A';
    const registeredAt = currentUserData?.createdAt ? formatDateValue(currentUserData.createdAt) : '-';
    const email = currentUserData?.email || currentUser?.email || 'N/A';
    const whatsapp = currentUserData?.whatsappNumber || currentUserData?.phone || '-';
    const location = currentUserData?.location || '-';
    const role = String(currentUserData?.role || 'payment');
    const status = String(currentUserData?.status || 'active');

    if (paymentUserName) paymentUserName.textContent = fullName;
    if (profileNameEl) profileNameEl.textContent = fullName;
    if (profileRegisteredAtEl) profileRegisteredAtEl.textContent = registeredAt;
    if (profileEmailEl) profileEmailEl.textContent = email;
    if (profileWhatsappEl) profileWhatsappEl.textContent = whatsapp;
    if (profileLocationEl) profileLocationEl.textContent = location;
    if (profileRoleEl) profileRoleEl.textContent = role.charAt(0).toUpperCase() + role.slice(1);
    if (profileStatusEl) profileStatusEl.textContent = status.charAt(0).toUpperCase() + status.slice(1);
}

function renderDashboardOverview() {
    const sentToPfaRecords = getSentToPfaRecords();
    const paidRecords = getPaidRecords();
    const clearedRecords = getClearedRecords();
    const allWorkflowRecords = [...sentToPfaRecords, ...paidRecords, ...clearedRecords];

    const sentGroups = buildAgentPaymentGroups(sentToPfaRecords);
    const paidGroups = buildAgentPaymentGroups(paidRecords);
    const clearedGroups = buildAgentPaymentGroups(clearedRecords);
    const totalCommission = allWorkflowRecords.reduce((sum, sub) => {
        const { commission2 } = getSubmissionFinancials(sub);
        return sum + commission2;
    }, 0);

    if (dashboardSentToPfaCount) dashboardSentToPfaCount.textContent = String(sentGroups.length);
    if (dashboardPaidCount) dashboardPaidCount.textContent = String(paidGroups.length);
    if (dashboardClearedCount) dashboardClearedCount.textContent = String(clearedGroups.length);
    if (dashboardTotalCommission) dashboardTotalCommission.textContent = formatCurrency(totalCommission);
}

function switchTab(tabId) {
    document.querySelectorAll('.nav-item').forEach((nav) => nav.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');
    const titles = {
        dashboard: 'Payment Dashboard',
        'sent-to-pfa': 'Sent to PFA',
        'paid-customers': 'Paid',
        'cleared-customers': 'Cleared',
        leave: 'Leave History',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Payment Dashboard';
    if (tabId === 'leave') {
        renderPaymentLeaveHistory().catch(() => {});
    }
}

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

function buildPaymentLeaveHistoryRecords(audits, mode = 'mine') {
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

async function loadPaymentLeaveHistory() {
    const auditSnap = await getDocs(query(collection(db, 'audit'), orderBy('timestamp', 'desc')));
    const audits = auditSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
    paymentMyLeaveHistory = buildPaymentLeaveHistoryRecords(audits, 'mine');
    paymentReliefLeaveHistory = buildPaymentLeaveHistoryRecords(audits, 'relief');
    paymentLeaveHistoryLoaded = true;
}

function renderPaymentLeaveRows(records, bodyId, includeOriginalUser = false) {
    const body = document.getElementById(bodyId);
    if (!body) return;
    if (!records.length) {
        body.innerHTML = `<tr><td colspan="${includeOriginalUser ? 8 : 7}" class="no-data">No leave records found</td></tr>`;
        return;
    }
    body.innerHTML = records.map((record) => `
        <tr>
            ${includeOriginalUser ? `<td>${record.originalUserEmail || '-'}</td>` : ''}
            <td>${formatDateValue(record.startAt)}</td>
            <td>${record.endAt ? formatDateValue(record.endAt) : '-'}</td>
            <td>${record.status}</td>
            <td>${record.relieverEmail || '-'}</td>
            <td>${record.movedCount}</td>
            <td>${record.returnedCount}/${record.finalizedCount}</td>
            <td><button class="action-btn" onclick="window.openPaymentLeaveApplications('${record.id}')"><i class="fas fa-eye"></i> View</button></td>
        </tr>
    `).join('');
}

async function renderPaymentLeaveHistory() {
    if (!paymentLeaveHistoryLoaded) {
        await loadPaymentLeaveHistory();
    }
    renderPaymentLeaveRows(paymentMyLeaveHistory, 'paymentMyLeaveTableBody', false);
    renderPaymentLeaveRows(paymentReliefLeaveHistory, 'paymentReliefLeaveTableBody', true);
}

window.openPaymentLeaveApplications = async (recordId) => {
    const record = [...paymentMyLeaveHistory, ...paymentReliefLeaveHistory].find((item) => item.id === recordId);
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
    const body = document.getElementById('paymentLeaveApplicationsBody');
    const title = document.getElementById('paymentLeaveApplicationsTitle');
    if (title) title.textContent = `Leave Applications - ${record.originalUserEmail || 'User'}`;
    if (body) {
        body.innerHTML = rows.length ? rows.map((sub) => `
            <tr>
                <td>${sub.customerName || 'Unknown'}</td>
                <td>${sub.status || '-'}</td>
                <td>${formatDateValue(sub.leaveCoverStartedAt)}</td>
                <td>${formatDateValue(sub.leaveCoverReturnedAt)}</td>
                <td>${formatDateValue(sub.leaveCoverFinalizedAt)}</td>
            </tr>
        `).join('') : '<tr><td colspan="5" class="no-data">No applications found for this leave record</td></tr>';
    }
    document.getElementById('paymentLeaveApplicationsModal')?.classList.add('active');
};

function getPaymentRecords() {
    return allSubmissions.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        if (status === 'sent_to_pfa' || status === 'rsa_submitted' || status === 'paid' || status === 'cleared') return true;
        // Backward compatibility for legacy final-submitted records
        return sub.finalSubmitted === true || sub.rsaSubmitted === true;
    });
}

function getSentToPfaRecords() {
    return getPaymentRecords().filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return status === 'sent_to_pfa' || status === 'rsa_submitted' || sub.finalSubmitted === true || sub.rsaSubmitted === true;
    });
}

function getPaidRecords() {
    return getPaymentRecords().filter((sub) => String(sub.status || '').toLowerCase() === 'paid');
}

function getClearedRecords() {
    return allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'cleared');
}

function renderPaymentQueue() {
    if (!paymentsTableBody) return;

    const paymentQueue = getSentToPfaRecords();

    if (paymentPendingCount) {
        paymentPendingCount.textContent = String(paymentQueue.length);
        paymentPendingCount.style.display = 'inline-block';
    }

    if (paymentQueue.length === 0) {
        paymentsTableBody.innerHTML = '<tr><td colspan="9" class="no-data">No applications sent to PFA yet</td></tr>';
        return;
    }

    paymentsTableBody.innerHTML = paymentQueue.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const queueDate = formatDateValue(sub?.rsaSubmittedAt || sub?.updatedAt);
        const uploaderLabel = getUploaderDisplayName(sub?.uploadedBy);
        const agentName = String(sub?.agentName || '').trim() || 'No Agent';
        const actionHtml = hasCommissionEligibleAgent(sub)
            ? `<button class="action-btn" style="background:#16a34a;color:#fff;border:none;" onclick="window.markSubmissionPaid('${sub.id}')"><i class="fas fa-check-circle"></i> Mark Paid</button>`
            : `<button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearSubmissionWithoutAgent('${sub.id}')"><i class="fas fa-check-double"></i> Clear</button>`;

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>
                    <strong>${escapeHtml(agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(String(sub?.agentAccountBank || '').trim() || '-')} • ${escapeHtml(String(sub?.agentAccountNumber || '').trim() || '-')}</div>
                </td>
                <td>${escapeHtml(uploaderLabel || '-')}</td>
                <td>${escapeHtml(pfa)}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td>${escapeHtml(queueDate)}</td>
                <td><span class="status-badge status-approved">Sent to PFA</span></td>
                <td>${actionHtml} <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');
}

function renderPaidCustomers() {
    if (!paidCustomersTableBody) return;

    const paidCustomers = getPaidRecords();
    const groupedPaid = buildAgentPaymentGroups(paidCustomers);
    if (paidCustomerCount) {
        paidCustomerCount.textContent = String(groupedPaid.length);
        paidCustomerCount.style.display = 'inline-block';
    }

    if (groupedPaid.length === 0) {
        paidCustomersTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No paid agent batches yet</td></tr>';
        return;
    }

    paidCustomersTableBody.innerHTML = groupedPaid.map((group) => {
        const breakdownId = `paid-breakdown-${toSafeDomId(group.key)}`;
        const paidDate = formatDateValue(group.latestPaidAt ? new Date(group.latestPaidAt) : null);
        return `
            <tr>
                <td>
                    <strong>${escapeHtml(group.agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</div>
                </td>
                <td>${escapeHtml(group.uploaderName || group.uploaderEmail || '-')}</td>
                <td>${group.customerCount}</td>
                <td>${formatCurrency(group.total25)}</td>
                <td>${formatCurrency(group.totalCommission)}</td>
                <td><span class="status-badge status-approved">Paid</span><div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(paidDate)}</div></td>
                <td><button class="action-btn agent-breakdown-toggle" onclick="window.togglePaymentAgentBreakdown('${breakdownId}', this)"><i class="fas fa-chevron-down"></i> Breakdown</button> <button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearPaidAgent('${group.key}')"><i class="fas fa-check-double"></i> Settle Agent</button></td>
            </tr>
            <tr id="${breakdownId}" class="agent-breakdown-row" style="display:none;">
                <td colspan="7">${renderAgentBreakdownTable(group, 'cleared')}</td>
            </tr>
        `;
    }).join('');
}

function renderClearedCustomers() {
    if (!clearedCustomersTableBody) return;

    const clearedCustomers = getClearedRecords();
    const groupedCleared = buildAgentPaymentGroups(clearedCustomers);
    if (clearedCustomerCount) {
        clearedCustomerCount.textContent = String(groupedCleared.length);
        clearedCustomerCount.style.display = 'inline-block';
    }

    if (groupedCleared.length === 0) {
        clearedCustomersTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No cleared agent batches yet</td></tr>';
        return;
    }

    clearedCustomersTableBody.innerHTML = groupedCleared.map((group) => {
        const breakdownId = `cleared-breakdown-${toSafeDomId(group.key)}`;
        const clearedDate = formatDateValue(group.latestClearedAt ? new Date(group.latestClearedAt) : null);
        return `
            <tr>
                <td>
                    <strong>${escapeHtml(group.agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</div>
                </td>
                <td>${escapeHtml(group.uploaderName || group.uploaderEmail || '-')}</td>
                <td>${group.customerCount}</td>
                <td>${formatCurrency(group.total25)}</td>
                <td>${formatCurrency(group.totalCommission)}</td>
                <td><span class="status-badge status-approved">Cleared</span><div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(clearedDate)}</div></td>
                <td><button class="action-btn agent-breakdown-toggle" onclick="window.togglePaymentAgentBreakdown('${breakdownId}', this)"><i class="fas fa-chevron-down"></i> Breakdown</button></td>
            </tr>
            <tr id="${breakdownId}" class="agent-breakdown-row" style="display:none;">
                <td colspan="7">${renderAgentBreakdownTable(group, 'paid')}</td>
            </tr>
        `;
    }).join('');
}

function loadSubmissions() {
    const q = query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc'));
    onSnapshot(q, async (snapshot) => {
        allSubmissions = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
        await primeUploaderNames(allSubmissions);
        renderDashboardOverview();
        renderPaymentQueue();
        renderPaidCustomers();
        renderClearedCustomers();
    }, () => {
        showNotification('Failed to load payment queue', 'error');
    });
}

window.markAgentPaid = async (groupKey) => {
    const agentItems = getPaymentRecords().filter((sub) => {
        const currentStatus = String(sub.status || '').toLowerCase();
        return getAgentPaymentKey(sub) === groupKey && (currentStatus === 'sent_to_pfa' || currentStatus === 'rsa_submitted' || sub.finalSubmitted === true || sub.rsaSubmitted === true);
    });
    if (!agentItems.length) {
        showNotification('Agent payment batch not found', 'error');
        return;
    }
    const group = buildAgentPaymentGroups(agentItems)[0];
    if (!group?.hasCommissionEligibleAgent) {
        showNotification('Applications without agent cannot be marked paid for commission.', 'warning');
        return;
    }
    const confirmed = confirm(`Mark ${group.agentName || 'this agent'} as PAID for ${group.customerCount} customer(s)?`);
    if (!confirmed) return;

    try {
        await Promise.all(agentItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentUser?.email || ''
        })));

        await addDoc(collection(db, 'audit'), {
            action: 'agent_commission_paid',
            agentKey: group.key,
            agentId: group.agentId || '',
            agentName: group.agentName || '',
            customerCount: group.customerCount,
            totalCommission: group.totalCommission,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        notifyStatusChangePush({
            currentUser,
            submissionId: agentItems[0].id,
            customerName: group.agentName || 'this agent',
            newStatus: 'paid',
            statusLabel: 'Paid',
            actionLabel: 'Agent Commission Marked Paid',
            message: `Commission for ${group.agentName || 'this agent'} was marked as paid for ${group.customerCount} customer(s).`
        }).catch(() => {});
        showNotification(`Marked ${group.agentName || 'agent'} as paid`, 'success');
    } catch (error) {
        showNotification('Failed to mark as paid', 'error');
    }
};

window.markSubmissionPaid = async (submissionId) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'error');
        return;
    }
    if (!hasCommissionEligibleAgent(sub)) {
        showNotification('This application has no agent commission to mark as paid.', 'warning');
        return;
    }

    const confirmed = confirm(`Mark ${sub.customerName || 'this application'} as paid?`);
    if (!confirmed) return;

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentUser?.email || ''
        });

        await addDoc(collection(db, 'audit'), {
            action: 'application_commission_paid',
            submissionId,
            customerName: sub.customerName || '',
            agentId: sub.agentId || '',
            agentName: sub.agentName || '',
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        notifyStatusChangePush({
            currentUser,
            submissionId,
            customerName: sub.customerName || '',
            newStatus: 'paid',
            statusLabel: 'Paid',
            actionLabel: 'Application Marked Paid',
            message: `Application for ${sub.customerName || 'this customer'} was marked as paid.`
        }).catch(() => {});

        showNotification(`Marked ${sub.customerName || 'application'} as paid`, 'success');
    } catch (error) {
        showNotification('Failed to mark application as paid', 'error');
    }
};

window.clearPaidAgent = async (groupKey) => {
    const paidItems = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid' && getAgentPaymentKey(s) === groupKey);
    if (paidItems.length === 0) {
        showNotification('No paid agent records to settle', 'info');
        return;
    }
    const group = buildAgentPaymentGroups(paidItems)[0];
    if (!group?.hasCommissionEligibleAgent) {
        showNotification('No-agent applications do not require commission settlement.', 'info');
        return;
    }
    const confirmed = confirm(`Settle commission for ${group.agentName || 'this agent'} across ${group.customerCount} customer(s)?`);
    if (!confirmed) return;

    try {
        await Promise.all(
            paidItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
                status: 'cleared',
                clearedAt: serverTimestamp(),
                clearedBy: currentUser?.email || ''
            }))
        );

        await addDoc(collection(db, 'audit'), {
            action: 'agent_commission_cleared',
            agentKey: group.key,
            agentId: group.agentId || '',
            agentName: group.agentName || '',
            count: paidItems.length,
            totalCommission: group.totalCommission,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        await Promise.all(
            paidItems.map((sub) => notifyStatusChangePush({
                currentUser,
                submissionId: sub.id,
                customerName: sub.customerName || '',
                newStatus: 'cleared',
                statusLabel: 'Cleared',
                actionLabel: 'Application Cleared',
                message: `Application for ${sub.customerName || 'this customer'} was cleared successfully.`
            }).catch(() => {}))
        );

        showNotification(`Settled ${group.agentName || 'agent'} commission`, 'success');
    } catch (error) {
        showNotification('Failed to clear paid records', 'error');
    }
};

window.clearPaidSubmissions = async () => {
    const paidItems = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid');
    if (!paidItems.length) {
        showNotification('No paid agent records to settle', 'info');
        return;
    }
    const groups = buildAgentPaymentGroups(paidItems).filter((group) => group.hasCommissionEligibleAgent);
    if (!groups.length) {
        showNotification('No paid commission batches to settle', 'info');
        return;
    }
    const confirmed = confirm(`Settle commission for ${groups.length} paid agent group(s)?`);
    if (!confirmed) return;
    try {
        for (const group of groups) {
            const groupItems = paidItems.filter((sub) => getAgentPaymentKey(sub) === group.key);
            await Promise.all(groupItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
                status: 'cleared',
                clearedAt: serverTimestamp(),
                clearedBy: currentUser?.email || ''
            })));
        }
        await addDoc(collection(db, 'audit'), {
            action: 'all_agent_commissions_cleared',
            count: paidItems.length,
            groupCount: groups.length,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification(`Settled ${groups.length} paid agent group(s)`, 'success');
    } catch (error) {
        showNotification('Failed to clear paid records', 'error');
    }
};

window.clearNoAgentApplications = async (groupKey) => {
    const items = getPaymentRecords().filter((sub) => getAgentPaymentKey(sub) === groupKey);
    if (!items.length) {
        showNotification('Application batch not found', 'error');
        return;
    }
    const group = buildAgentPaymentGroups(items)[0];
    const confirmed = confirm(`Clear ${group.customerCount} application(s) without agent commission?`);
    if (!confirmed) return;

    try {
        await Promise.all(items.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
            status: 'cleared',
            clearedAt: serverTimestamp(),
            clearedBy: currentUser?.email || '',
            clearedWithoutAgentCommission: true
        })));
        await addDoc(collection(db, 'audit'), {
            action: 'applications_cleared_without_agent_commission',
            agentKey: group.key,
            count: items.length,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification(`Cleared ${group.customerCount} no-agent application(s)`, 'success');
    } catch (error) {
        showNotification('Failed to clear no-agent applications', 'error');
    }
};

window.clearSubmissionWithoutAgent = async (submissionId) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'error');
        return;
    }

    const confirmed = confirm(`Clear ${sub.customerName || 'this application'} without agent commission?`);
    if (!confirmed) return;

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'cleared',
            clearedAt: serverTimestamp(),
            clearedBy: currentUser?.email || '',
            clearedWithoutAgentCommission: true
        });

        await addDoc(collection(db, 'audit'), {
            action: 'application_cleared_without_agent_commission',
            submissionId,
            customerName: sub.customerName || '',
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification(`Cleared ${sub.customerName || 'application'}`, 'success');
    } catch (error) {
        showNotification('Failed to clear application', 'error');
    }
};

window.togglePaymentAgentBreakdown = (rowId, btn) => {
    const row = document.getElementById(rowId);
    if (!row) return;
    const isOpen = row.style.display !== 'none';
    row.style.display = isOpen ? 'none' : 'table-row';
    if (btn) {
        const icon = btn.querySelector('i');
        if (icon) icon.className = isOpen ? 'fas fa-chevron-down' : 'fas fa-chevron-up';
        const textNode = Array.from(btn.childNodes).find((node) => node.nodeType === Node.TEXT_NODE);
        if (textNode) textNode.textContent = isOpen ? ' Breakdown' : ' Hide';
    }
};

window.signOutUser = async () => {
    try { await signOut(auth); } catch (e) {}
    window.location.href = 'index.html';
};

function forceHardRefresh() {
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
}

document.getElementById('forceRefreshBtn')?.addEventListener('click', (e) => {
    e.preventDefault();
    forceHardRefresh();
});
document.getElementById('paymentLeaveMineBtn')?.addEventListener('click', () => {
    document.getElementById('paymentLeaveMineSection')?.style.setProperty('display', '');
    document.getElementById('paymentLeaveReliefSection')?.style.setProperty('display', 'none');
});
document.getElementById('paymentLeaveReliefBtn')?.addEventListener('click', () => {
    document.getElementById('paymentLeaveMineSection')?.style.setProperty('display', 'none');
    document.getElementById('paymentLeaveReliefSection')?.style.setProperty('display', '');
});
document.getElementById('closePaymentLeaveApplicationsBtn')?.addEventListener('click', () => {
    document.getElementById('paymentLeaveApplicationsModal')?.classList.remove('active');
});
window.addEventListener('click', (e) => {
    if (e.target === document.getElementById('paymentLeaveApplicationsModal')) {
        document.getElementById('paymentLeaveApplicationsModal')?.classList.remove('active');
    }
});

document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        switchTab(item.dataset.tab);
    });
});

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
        if (role === 'admin') {
            window.location.href = 'admin-dashboard.html';
            return;
        }
        if (role !== 'payment') {
            window.location.href = 'index.html';
            return;
        }

        currentUserData = userData;
        renderProfile();
        loadSubmissions();
    } catch (error) {
        showNotification('Could not validate session', 'error');
        window.location.href = 'index.html';
    }
});
