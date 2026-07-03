import { auth, db } from './firebase-config.js?v=20260625c';
import { performAppLogout } from './shared/logout.js?v=20260625b';
import {
    collection,
    addDoc,
    doc,
    onSnapshot,
    serverTimestamp,
    arrayUnion,
    updateDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { formatAppDateTime } from './shared/app-time.js';
import { getCurrentUserProfile as getCurrentUserProfileShared } from './shared/user-directory.js?v=20260518a';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionCurrentStageEntryAt
} from './shared/submission-stage.js?v=20260609a';
import {
    getSubmissionCommissionAmount,
    resolveSubmissionCommissionRate
} from './shared/commission-config.js?v=20260507a';
import { notifyUserPushEvent } from './push-alerts.js';

let currentUser = null;
let currentUserData = null;
let allUsers = [];
let allSubmissions = [];
let currentTab = 'overview';
let currentAuditPaidScope = 'mine';
let currentPaymentReport = null;
let pendingPaymentReportRequest = { kind: 'paid' };
let currentPaymentPdfPreviewUrl = '';
let currentPaymentPdfPreviewBlob = null;
let currentPaymentPdfPreviewFileName = 'payment-report.pdf';
let usersListenerStarted = false;
let submissionsListenerStarted = false;
const PAYMENT_RATE_CUTOFF_MS = new Date('2026-05-07T00:00:00+01:00').getTime();

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
const auditSentToPfaTableBody = document.getElementById('auditSentToPfaTableBody');
const auditOverviewPendingTableBody = document.getElementById('auditOverviewPendingTableBody');
const auditPaidTableBody = document.getElementById('auditPaidTableBody');
const auditClearedTableBody = document.getElementById('auditClearedTableBody');
const auditRejectedTableBody = document.getElementById('auditRejectedTableBody');
const exportAuditPendingReportBtn = document.getElementById('exportAuditPendingReportBtn');
const exportAuditPaidReportBtn = document.getElementById('exportAuditPaidReportBtn');
const exportAuditClearedReportBtn = document.getElementById('exportAuditClearedReportBtn');
const exportAuditRejectedReportBtn = document.getElementById('exportAuditRejectedReportBtn');
const paymentReportRangeModal = document.getElementById('paymentReportRangeModal');
const paymentReportPreviewModal = document.getElementById('paymentReportPreviewModal');
const paymentReportStartDate = document.getElementById('paymentReportStartDate');
const paymentReportEndDate = document.getElementById('paymentReportEndDate');
const paymentReportPreviewMeta = document.getElementById('paymentReportPreviewMeta');
const paymentReportSummaryChips = document.getElementById('paymentReportSummaryChips');
const paymentReportSummaryBody = document.getElementById('paymentReportSummaryBody');
const paymentReportDetailsBody = document.getElementById('paymentReportDetailsBody');
const generatePaymentReportBtn = document.getElementById('generatePaymentReportBtn');
const exportPaymentReportExcelBtn = document.getElementById('exportPaymentReportExcelBtn');
const exportPaymentReportPdfBtn = document.getElementById('exportPaymentReportPdfBtn');
const paymentPdfPreviewModal = document.getElementById('paymentPdfPreviewModal');
const paymentPdfPreviewFrame = document.getElementById('paymentPdfPreviewFrame');
const downloadPaymentPdfPreviewBtn = document.getElementById('downloadPaymentPdfPreviewBtn');

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

function getTimestampMillis(value) {
    return getStageTimestampMillis(value);
}

function formatDateOnly(value) {
    const ms = getTimestampMillis(value);
    if (!ms) return '-';
    try {
        return new Date(ms).toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: 'numeric'
        });
    } catch (_) {
        return new Date(ms).toISOString().slice(0, 10);
    }
}

function formatInputDateValue(value) {
    const ms = getTimestampMillis(value);
    if (!ms) return '';
    const d = new Date(ms);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (_) {
        return `NGN ${num.toLocaleString(undefined, { maximumFractionDigits: 2 })}`;
    }
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

function getSubmissionFinancials(sub = {}) {
    const details = sub.customerDetails || {};
    const rsaBalance = parseMoney(details.rsaBalance || sub.rsaBalance || 0);
    const stored25 = parseMoney(details.rsa25Percent || sub.rsa25Percent || 0);
    const twentyFive = stored25 || roundDownToNearestThousand(rsaBalance * 0.25);
    const commission = getSubmissionCommissionAmount(sub, twentyFive);
    const commissionRate = resolveSubmissionCommissionRate(sub);
    return { rsaBalance, twentyFive, commission, commissionRate };
}

function getSubmissionPfaName(sub = {}) {
    return String(sub?.customerDetails?.pfa || sub?.pfa || '').trim() || '-';
}

function getSubmissionRatePercent(sub = {}) {
    const uploadedAtMs = getTimestampMillis(sub?.uploadedAt || sub?.createdAt || sub?.updatedAt);
    if (uploadedAtMs && uploadedAtMs < PAYMENT_RATE_CUTOFF_MS) return 10;
    return 7;
}

function getSubmissionRateLabel(sub = {}) {
    return `${getSubmissionRatePercent(sub)}%`;
}

function getPaymentStatusLabel(sub = {}) {
    const status = String(sub?.status || '').toLowerCase();
    if (status === 'cleared') return 'Cleared';
    if (status === 'paid') return 'Paid';
    return 'Sent to PFA';
}

function isPaymentSubmissionAttended(sub = {}) {
    const status = String(sub?.status || '').toLowerCase();
    return status === 'paid' || status === 'cleared';
}

function getRangeBoundaryMs(dateValue, mode = 'start') {
    if (!dateValue) return 0;
    const suffix = mode === 'end' ? 'T23:59:59.999+01:00' : 'T00:00:00+01:00';
    return new Date(`${dateValue}${suffix}`).getTime();
}

function getCustomerAccountNumber(sub = {}) {
    return String(
        sub?.customerDetails?.accountNo ||
        sub?.customerDetails?.accountNumber ||
        sub?.accountNo ||
        sub?.accountNumber ||
        '-'
    ).trim() || '-';
}

function renderCopyableCustomerAccount(sub = {}) {
    const accountNumber = getCustomerAccountNumber(sub);
    if (!accountNumber || accountNumber === '-') return '-';
    return `
        <span class="copy-account-cell">
            <span class="copy-account-value">${escapeHtml(accountNumber)}</span>
            <button type="button" class="copy-account-btn" data-copy-account="${escapeHtml(accountNumber)}" title="Copy account number" aria-label="Copy account number">
                <i class="fas fa-copy"></i>
            </button>
        </span>
    `;
}

async function copyTextToClipboard(text = '') {
    const value = String(text || '').trim();
    if (!value) return false;
    if (navigator.clipboard?.writeText && window.isSecureContext) {
        await navigator.clipboard.writeText(value);
        return true;
    }
    const textarea = document.createElement('textarea');
    textarea.value = value;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.left = '-9999px';
    document.body.appendChild(textarea);
    textarea.select();
    let copied = false;
    try {
        copied = document.execCommand('copy');
    } finally {
        textarea.remove();
    }
    return copied;
}

function getUserDisplayName(email = '') {
    const normalized = normalizeEmail(email);
    if (!normalized) return '-';
    const user = allUsers.find((entry) => normalizeEmail(entry.email) === normalized);
    return user?.fullName || email;
}

function roleLabel(role) {
    const normalized = String(role || '').trim().toLowerCase();
    if (normalized === 'super_admin') return 'Super Admin';
    if (normalized === 'admin') return 'Admin';
    if (normalized === 'reports_monitoring') return 'Audit';
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

function isSentToPfaLifecycle(sub = {}) {
    const status = String(sub.status || '').toLowerCase();
    return (
        status === 'sent_to_pfa' ||
        status === 'rsa_submitted' ||
        (
            (sub.finalSubmitted === true || sub.rsaSubmitted === true) &&
            status !== 'paid' &&
            status !== 'cleared'
        )
    );
}

function getAuditSentToPfaRows() {
    return allSubmissions
        .filter((sub) => isSentToPfaLifecycle(sub))
        .filter((sub) => sub.paymentMadeByUploader === true && String(sub.auditCommissionStatus || '').toLowerCase() === 'pending')
        .sort((a, b) => getStageTimestampMillis(b.auditCommissionSubmittedAt || b.paymentMadeAt || getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(a.auditCommissionSubmittedAt || a.paymentMadeAt || getSubmissionCurrentStageEntryAt(a)));
}

function getAuditApprovalEmail(sub = {}) {
    return normalizeEmail(sub.auditCommissionAcceptedBy || sub.paidBy || sub.commissionPaidBy || '');
}

function getAuditPaidRows(scope = currentAuditPaidScope) {
    const currentEmail = normalizeEmail(currentUser?.email);
    return allSubmissions
        .filter((sub) => String(sub.status || '').toLowerCase() === 'paid')
        .filter((sub) => {
            if (scope === 'all') return true;
            const approvedBy = getAuditApprovalEmail(sub);
            if (!currentEmail) return scope !== 'mine';
            return scope === 'mine' ? approvedBy === currentEmail : approvedBy !== currentEmail;
        })
        .sort((a, b) => getStageTimestampMillis(b.paidAt || getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(a.paidAt || getSubmissionCurrentStageEntryAt(a)));
}

function getAuditClearedRows() {
    return allSubmissions
        .filter((sub) => String(sub.status || '').toLowerCase() === 'cleared')
        .sort((a, b) => getStageTimestampMillis(b.clearedAt || getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(a.clearedAt || getSubmissionCurrentStageEntryAt(a)));
}

function getAuditRejectedRows() {
    const currentEmail = normalizeEmail(currentUser?.email);
    return allSubmissions
        .filter((sub) => String(sub.auditCommissionStatus || '').toLowerCase() === 'rejected')
        .filter((sub) => normalizeEmail(sub.auditCommissionRejectedBy || '') === currentEmail)
        .sort((a, b) => getStageTimestampMillis(b.auditCommissionRejectedAt) - getStageTimestampMillis(a.auditCommissionRejectedAt));
}

function setCountBadge(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = String(value);
}

function renderAuditWorkflowBadges() {
    setCountBadge('auditSentToPfaCountBadge', getAuditSentToPfaRows().length);
    setCountBadge('auditPaidCountBadge', getAuditPaidRows('all').length);
    setCountBadge('auditClearedCountBadge', getAuditClearedRows().length);
    setCountBadge('auditRejectedCountBadge', getAuditRejectedRows().length);
}

function renderOverview() {
    const awaitingAuditCount = getAuditSentToPfaRows().length;
    const paidCount = getAuditPaidRows('all').length;
    const clearedCount = getAuditClearedRows().length;

    setCountBadge('overviewUsersCount', awaitingAuditCount);
    setCountBadge('overviewSentCount', paidCount);
    setCountBadge('overviewPaidClearedCount', clearedCount);
    renderAuditWorkflowBadges();
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

function renderAuditMoneyRows(body, rows, mode) {
    if (!body) return;
    if (!rows.length) {
        const label = mode === 'sent' ? 'pending payment requests' : `${mode} records`;
        const colspan = mode === 'sent' ? 7 : 9;
        body.innerHTML = `<tr><td colspan="${colspan}" class="no-data">No ${escapeHtml(label)}</td></tr>`;
        return;
    }

    body.innerHTML = rows.map((sub) => {
        const { rsaBalance, twentyFive, commission } = getSubmissionFinancials(sub);
        const uploaderName = getUserDisplayName(sub.uploadedBy || sub.auditCommissionSubmittedBy || '');
        const actionCell = mode === 'sent'
            ? `<button class="action-btn" onclick="window.acceptAuditCommission('${sub.id}')"><i class="fas fa-check"></i> Accept</button>
               <button class="action-btn" style="background:#b91c1c;color:#fff;border:none;" onclick="window.rejectAuditCommission('${sub.id}')"><i class="fas fa-times"></i> Reject</button>`
            : `<button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>`;
        const officerEmail = mode === 'paid'
            ? (sub.auditCommissionAcceptedBy || sub.paidBy || sub.commissionPaidBy || '')
            : mode === 'cleared'
                ? (sub.clearedBy || '')
                : '';
        const officerCell = mode === 'paid' || mode === 'cleared'
            ? `<td>${escapeHtml(getUserDisplayName(officerEmail))}</td>`
            : '';
        const timeCell = mode === 'paid'
            ? `<td>${escapeHtml(formatDate(sub.paidAt))}</td>`
            : mode === 'cleared'
                ? `<td>${escapeHtml(formatDate(sub.clearedAt))}</td>`
                : `<td>${actionCell}</td>`;

        if (mode === 'sent') {
            return `
                <tr>
                    <td>${escapeHtml(uploaderName)}</td>
                    <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                    <td>${formatCurrency(rsaBalance)}</td>
                    <td>${formatCurrency(twentyFive)}</td>
                    <td><strong>${formatCurrency(commission)}</strong></td>
                    <td>${renderCopyableCustomerAccount(sub)}</td>
                    <td>${actionCell}</td>
                </tr>
            `;
        }

        return `
            <tr>
                <td>${escapeHtml(uploaderName)}</td>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${formatCurrency(rsaBalance)}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td><strong>${formatCurrency(commission)}</strong></td>
                <td>${renderCopyableCustomerAccount(sub)}</td>
                ${officerCell}
                ${timeCell}
                <td>${actionCell}</td>
            </tr>
        `;
    }).join('');
}

function renderAuditRejectedRows(body, rows) {
    if (!body) return;
    if (!rows.length) {
        body.innerHTML = `<tr><td colspan="10" class="no-data">No rejected commission requests</td></tr>`;
        return;
    }

    body.innerHTML = rows.map((sub) => {
        const { rsaBalance, twentyFive, commission } = getSubmissionFinancials(sub);
        const uploaderName = getUserDisplayName(sub.uploadedBy || sub.auditCommissionSubmittedBy || '');
        const rejectionReason = String(sub.auditCommissionRejectionReason || 'No reason provided').trim();
        const actionCell = `<button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>`;

        return `
            <tr>
                <td>${escapeHtml(uploaderName)}</td>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${formatCurrency(rsaBalance)}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td><strong>${formatCurrency(commission)}</strong></td>
                <td>${renderCopyableCustomerAccount(sub)}</td>
                <td>${escapeHtml(getUserDisplayName(sub.auditCommissionRejectedBy || ''))}</td>
                <td>${escapeHtml(formatDate(sub.auditCommissionRejectedAt))}</td>
                <td><strong>${escapeHtml(rejectionReason)}</strong></td>
                <td>${actionCell}</td>
            </tr>
        `;
    }).join('');
}

function renderAuditWorkflowTabs() {
    renderAuditWorkflowBadges();
    document.querySelectorAll('[data-audit-paid-scope]').forEach((button) => {
        button.classList.toggle('active', button.dataset.auditPaidScope === currentAuditPaidScope);
    });
    renderAuditMoneyRows(auditOverviewPendingTableBody, getAuditSentToPfaRows(), 'sent');
    renderAuditMoneyRows(auditSentToPfaTableBody, getAuditSentToPfaRows(), 'sent');
    renderAuditMoneyRows(auditPaidTableBody, getAuditPaidRows(currentAuditPaidScope), 'paid');
    renderAuditMoneyRows(auditClearedTableBody, getAuditClearedRows(), 'cleared');
    renderAuditRejectedRows(auditRejectedTableBody, getAuditRejectedRows());
}

function buildPaymentStageReport(records = [], options = {}) {
    const {
        title = 'Payment Report',
        exportKey = 'payment-report',
        dateLabel = 'Date',
        getDateMs = (sub) => getTimestampMillis(sub.paidAt || getSubmissionCurrentStageEntryAt(sub)),
        statusResolver = (sub) => getPaymentStatusLabel(sub),
        attendedResolver = (sub) => isPaymentSubmissionAttended(sub)
    } = options;

    const summaryMap = new Map();
    const details = records
        .map((sub) => {
            const { twentyFive, commission } = getSubmissionFinancials(sub);
            const dateMs = getDateMs(sub);
            const attended = attendedResolver(sub);
            const statusText = statusResolver(sub);
            const detailRow = {
                id: sub.id,
                receivedAtMs: dateMs,
                receivedDate: formatDateOnly(dateMs),
                customerName: String(sub?.customerName || 'Unknown').trim() || 'Unknown',
                agentName: String(sub?.agentName || '').trim() || 'No Agent',
                uploaderName: getUserDisplayName(sub?.uploadedBy),
                pfa: getSubmissionPfaName(sub),
                twentyFive,
                commission,
                rateLabel: getSubmissionRateLabel(sub),
                statusLabel: statusText,
                attendedLabel: attended ? 'Yes' : 'No',
                paidDate: formatDate(sub?.paidAt),
                clearedDate: formatDate(sub?.clearedAt)
            };

            const summaryKey = formatInputDateValue(dateMs);
            const existingSummary = summaryMap.get(summaryKey) || {
                key: summaryKey,
                dateLabel: formatDateOnly(dateMs),
                received: 0,
                attended: 0,
                pending: 0
            };
            existingSummary.received += 1;
            if (attended) existingSummary.attended += 1;
            else existingSummary.pending += 1;
            summaryMap.set(summaryKey, existingSummary);
            return detailRow;
        })
        .sort((a, b) => b.receivedAtMs - a.receivedAtMs);

    const summaryRows = Array.from(summaryMap.values()).sort((a, b) => String(a.key).localeCompare(String(b.key)));
    const totals = summaryRows.reduce((acc, row) => ({
        received: acc.received + row.received,
        attended: acc.attended + row.attended,
        pending: acc.pending + row.pending
    }), { received: 0, attended: 0, pending: 0 });

    return {
        title,
        exportKey,
        rangeStart: options.rangeStart || '',
        rangeEnd: options.rangeEnd || '',
        metaText: `${title} generated from ${records.length} record(s). Summary grouped by ${dateLabel.toLowerCase()}.`,
        summaryRows,
        details,
        totals,
        dateCount: summaryRows.length
    };
}

function createStageReportFromDateRange(request, startDate, endDate) {
    const startMs = getRangeBoundaryMs(startDate, 'start');
    const endMs = getRangeBoundaryMs(endDate, 'end');

    if (request?.kind === 'sent-to-pfa') {
        const records = getAuditSentToPfaRows().filter((sub) => {
            const dateMs = getTimestampMillis(sub.auditCommissionSubmittedAt || sub.paymentMadeAt || getSubmissionCurrentStageEntryAt(sub));
            return dateMs >= startMs && dateMs <= endMs;
        });
        return buildPaymentStageReport(records, {
            title: 'Sent to PFA Report',
            exportKey: `sent-to-pfa-report-${startDate}-to-${endDate}`,
            rangeStart: startDate,
            rangeEnd: endDate,
            dateLabel: 'sent date',
            getDateMs: (sub) => getTimestampMillis(sub.auditCommissionSubmittedAt || sub.paymentMadeAt || getSubmissionCurrentStageEntryAt(sub)),
            statusResolver: () => 'Sent to PFA',
            attendedResolver: () => false
        });
    }

    if (request?.kind === 'cleared') {
        const records = getAuditClearedRows().filter((sub) => {
            const dateMs = getTimestampMillis(sub.clearedAt);
            return dateMs >= startMs && dateMs <= endMs;
        });
        return buildPaymentStageReport(records, {
            title: 'Cleared Report',
            exportKey: `cleared-report-${startDate}-to-${endDate}`,
            rangeStart: startDate,
            rangeEnd: endDate,
            dateLabel: 'cleared date',
            getDateMs: (sub) => getTimestampMillis(sub.clearedAt),
            statusResolver: () => 'Cleared',
            attendedResolver: () => true
        });
    }

    if (request?.kind === 'rejected') {
        const records = getAuditRejectedRows().filter((sub) => {
            const dateMs = getTimestampMillis(sub.auditCommissionRejectedAt);
            return dateMs >= startMs && dateMs <= endMs;
        });
        return buildPaymentStageReport(records, {
            title: 'Rejected Report',
            exportKey: `rejected-report-${startDate}-to-${endDate}`,
            rangeStart: startDate,
            rangeEnd: endDate,
            dateLabel: 'rejected date',
            getDateMs: (sub) => getTimestampMillis(sub.auditCommissionRejectedAt),
            statusResolver: () => 'Rejected',
            attendedResolver: () => false
        });
    }

    const records = getAuditPaidRows(currentAuditPaidScope).filter((sub) => {
        const dateMs = getTimestampMillis(sub.paidAt);
        return dateMs >= startMs && dateMs <= endMs;
    });
    return buildPaymentStageReport(records, {
        title: 'Paid Report',
        exportKey: `paid-report-${startDate}-to-${endDate}`,
        rangeStart: startDate,
        rangeEnd: endDate,
        dateLabel: 'paid date',
        getDateMs: (sub) => getTimestampMillis(sub.paidAt),
        statusResolver: () => 'Paid',
        attendedResolver: () => true
    });
}

function renderPaymentReportPreview(report) {
    currentPaymentReport = report;
    if (paymentReportPreviewMeta) paymentReportPreviewMeta.textContent = report.metaText || 'Payment report preview.';
    if (paymentReportSummaryChips) {
        paymentReportSummaryChips.innerHTML = [
            { label: 'Days Covered', value: report.dateCount, bg: '#dbeafe', color: '#1d4ed8' },
            { label: 'Applications Received', value: report.totals.received, bg: '#e2e8f0', color: '#334155' },
            { label: 'Attended', value: report.totals.attended, bg: '#dcfce7', color: '#166534' },
            { label: 'Pending', value: report.totals.pending, bg: '#fee2e2', color: '#b91c1c' }
        ].map((item) => `<span style="display:inline-flex;align-items:center;gap:6px;padding:8px 12px;border-radius:999px;background:${item.bg};color:${item.color};font-size:12px;font-weight:700;"><span>${escapeHtml(item.label)}</span><span>${escapeHtml(String(item.value))}</span></span>`).join('');
    }
    if (paymentReportSummaryBody) {
        paymentReportSummaryBody.innerHTML = report.summaryRows.length
            ? report.summaryRows.map((row) => `<tr><td>${escapeHtml(row.dateLabel)}</td><td>${escapeHtml(String(row.received))}</td><td>${escapeHtml(String(row.attended))}</td><td>${escapeHtml(String(row.pending))}</td></tr>`).join('')
            : '<tr><td colspan="4" class="no-data">No payment applications found for this date range</td></tr>';
    }
    if (paymentReportDetailsBody) {
        paymentReportDetailsBody.innerHTML = report.details.length
            ? report.details.map((row) => `
                <tr>
                    <td>${escapeHtml(row.receivedDate)}</td>
                    <td><strong>${escapeHtml(row.customerName)}</strong></td>
                    <td>${escapeHtml(row.agentName)}</td>
                    <td>${escapeHtml(row.uploaderName || '-')}</td>
                    <td>${escapeHtml(row.pfa)}</td>
                    <td>${formatCurrency(row.twentyFive)}</td>
                    <td>${escapeHtml(row.rateLabel)}</td>
                    <td>${formatCurrency(row.commission)}</td>
                    <td>${escapeHtml(row.statusLabel)}</td>
                    <td>${escapeHtml(row.attendedLabel)}</td>
                    <td>${escapeHtml(row.paidDate)}</td>
                    <td>${escapeHtml(row.clearedDate)}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="12" class="no-data">No application breakdown available for this date range</td></tr>';
    }
}

function saveBlob(fileName, blob) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName.replace(/[\\/:*?"<>|]/g, '_');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function openPaymentPdfPreview(blob, fileName = 'payment-report.pdf') {
    if (!blob) return;
    if (!paymentPdfPreviewModal || !paymentPdfPreviewFrame) {
        saveBlob(fileName, blob);
        return;
    }
    if (currentPaymentPdfPreviewUrl) URL.revokeObjectURL(currentPaymentPdfPreviewUrl);
    currentPaymentPdfPreviewBlob = blob;
    currentPaymentPdfPreviewFileName = fileName.replace(/[\\/:*?"<>|]/g, '_');
    currentPaymentPdfPreviewUrl = URL.createObjectURL(blob);
    paymentPdfPreviewFrame.src = currentPaymentPdfPreviewUrl;
    paymentPdfPreviewModal.classList.add('active');
}

function closePaymentPdfPreviewModal() {
    paymentPdfPreviewModal?.classList.remove('active');
    if (paymentPdfPreviewFrame) paymentPdfPreviewFrame.src = '';
    if (currentPaymentPdfPreviewUrl) URL.revokeObjectURL(currentPaymentPdfPreviewUrl);
    currentPaymentPdfPreviewUrl = '';
    currentPaymentPdfPreviewBlob = null;
}

async function exportPaymentReportExcel(report) {
    if (!report) return;
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Payment Dashboard';
    workbook.created = new Date();

    const summarySheet = workbook.addWorksheet('Daily Summary');
    summarySheet.columns = [
        { header: 'Date', key: 'dateLabel', width: 18 },
        { header: 'Applications Received', key: 'received', width: 22 },
        { header: 'Attended', key: 'attended', width: 14 },
        { header: 'Pending', key: 'pending', width: 14 }
    ];
    summarySheet.addRow({ dateLabel: 'Range', received: `${report.rangeStart} to ${report.rangeEnd}`, attended: '', pending: '' });
    summarySheet.addRow({});
    report.summaryRows.forEach((row) => summarySheet.addRow(row));
    summarySheet.addRow({});
    summarySheet.addRow({ dateLabel: 'Total', received: report.totals.received, attended: report.totals.attended, pending: report.totals.pending });
    summarySheet.getRow(1).font = { bold: true };
    summarySheet.getRow(3).font = { bold: true };

    const detailsSheet = workbook.addWorksheet('Application Breakdown');
    detailsSheet.columns = [
        { header: 'Received Date', key: 'receivedDate', width: 18 },
        { header: 'Customer', key: 'customerName', width: 28 },
        { header: 'Agent', key: 'agentName', width: 24 },
        { header: 'Uploader', key: 'uploaderName', width: 24 },
        { header: 'PFA', key: 'pfa', width: 20 },
        { header: '25% Balance', key: 'twentyFive', width: 18, style: { numFmt: '#,##0.00' } },
        { header: 'Rate', key: 'rateLabel', width: 12 },
        { header: 'Commission', key: 'commission', width: 18, style: { numFmt: '#,##0.00' } },
        { header: 'Status', key: 'statusLabel', width: 16 },
        { header: 'Attended', key: 'attendedLabel', width: 12 },
        { header: 'Paid Date', key: 'paidDate', width: 22 },
        { header: 'Cleared Date', key: 'clearedDate', width: 22 }
    ];
    report.details.forEach((row) => detailsSheet.addRow(row));
    detailsSheet.getRow(1).font = { bold: true };

    const buffer = await workbook.xlsx.writeBuffer();
    saveBlob(`${report.exportKey || 'payment-report'}.xlsx`, new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
}

async function exportPaymentReportPdf(report) {
    if (!report || !window.jspdf?.jsPDF) throw new Error('PDF library not available.');
    const pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    pdf.setFontSize(16);
    pdf.text(report.title || 'Payment Report', 40, 36);
    pdf.setFontSize(10);
    pdf.text(report.metaText || '', 40, 54);
    pdf.text(`Received: ${report.totals.received}   Attended: ${report.totals.attended}   Pending: ${report.totals.pending}`, 40, 70);
    pdf.autoTable({
        startY: 88,
        head: [['Date', 'Received', 'Attended', 'Pending']],
        body: report.summaryRows.map((row) => [row.dateLabel, row.received, row.attended, row.pending]),
        styles: { fontSize: 9 },
        headStyles: { fillColor: [15, 59, 103] }
    });
    pdf.autoTable({
        startY: pdf.lastAutoTable.finalY + 18,
        head: [['Received Date', 'Customer', 'Agent', 'Uploader', 'PFA', '25% Balance', 'Rate', 'Commission', 'Status', 'Attended']],
        body: report.details.map((row) => [
            row.receivedDate,
            row.customerName,
            row.agentName,
            row.uploaderName,
            row.pfa,
            Number(row.twentyFive || 0).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
            row.rateLabel,
            Number(row.commission || 0).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
            row.statusLabel,
            row.attendedLabel
        ]),
        styles: { fontSize: 7, cellPadding: 2.5, overflow: 'linebreak', valign: 'middle' },
        headStyles: { fillColor: [15, 118, 110] },
        showHead: 'everyPage',
        margin: { left: 18, right: 18 },
        tableWidth: 'auto',
        columnStyles: {
            0: { cellWidth: 58 },
            1: { cellWidth: 112 },
            2: { cellWidth: 88 },
            3: { cellWidth: 88 },
            4: { cellWidth: 58 },
            5: { cellWidth: 72, halign: 'right' },
            6: { cellWidth: 36, halign: 'center' },
            7: { cellWidth: 72, halign: 'right' },
            8: { cellWidth: 54, halign: 'center' },
            9: { cellWidth: 48, halign: 'center' }
        }
    });
    openPaymentPdfPreview(pdf.output('blob'), `${report.exportKey || 'payment-report'}.pdf`);
}

function resetPaymentReportPreviewTabs() {
    document.querySelectorAll('.payment-report-preview-tab').forEach((btn) => {
        const isSummary = btn.dataset.paymentReportView === 'summary';
        btn.classList.toggle('active', isSummary);
        btn.style.background = isSummary ? '#0f3b67' : '';
        btn.style.color = isSummary ? '#fff' : '';
        btn.style.border = isSummary ? 'none' : '';
    });
    document.getElementById('paymentReportSummaryView')?.style.setProperty('display', '');
    document.getElementById('paymentReportDetailsView')?.style.setProperty('display', 'none');
}

function showPaymentReportPreviewTab(viewName) {
    document.querySelectorAll('.payment-report-preview-tab').forEach((btn) => {
        const isActive = btn.dataset.paymentReportView === viewName;
        btn.classList.toggle('active', isActive);
        btn.style.background = isActive ? '#0f3b67' : '';
        btn.style.color = isActive ? '#fff' : '';
        btn.style.border = isActive ? 'none' : '';
    });
    document.getElementById('paymentReportSummaryView')?.style.setProperty('display', viewName === 'summary' ? '' : 'none');
    document.getElementById('paymentReportDetailsView')?.style.setProperty('display', viewName === 'details' ? '' : 'none');
}

function openPaymentReportRangeModal() {
    const today = new Date();
    const sixDaysAgo = new Date(today.getTime() - (6 * 24 * 60 * 60 * 1000));
    if (paymentReportStartDate && !paymentReportStartDate.value) paymentReportStartDate.value = formatInputDateValue(sixDaysAgo);
    if (paymentReportEndDate && !paymentReportEndDate.value) paymentReportEndDate.value = formatInputDateValue(today);
    paymentReportRangeModal?.classList.add('active');
}

function closePaymentReportRangeModal() {
    paymentReportRangeModal?.classList.remove('active');
}

function openPaymentReportPreviewModal() {
    resetPaymentReportPreviewTabs();
    paymentReportPreviewModal?.classList.add('active');
}

function closePaymentReportPreviewModal() {
    paymentReportPreviewModal?.classList.remove('active');
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
    renderAuditWorkflowTabs();
}

function switchTab(tabId) {
    currentTab = tabId;
    ensureDataForTab(tabId);
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');

    const titles = {
        overview: 'Audit Overview',
        'sent-to-pfa': 'Pending Request',
        paid: 'Paid',
        cleared: 'Cleared',
        rejected: 'Rejected',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Audit';
}

function ensureDataForTab(tabId) {
    if (tabId === 'overview' || tabId === 'sent-to-pfa' || tabId === 'paid' || tabId === 'cleared') loadUsers();
    if (tabId === 'overview' || tabId === 'sent-to-pfa' || tabId === 'paid' || tabId === 'cleared') loadSubmissions();
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
        ['Customer Account Number', getCustomerAccountNumber(sub)],
        ['Audit Commission Status', statusLabel(sub.auditCommissionStatus || '-')],
        ['Audit Rejection Reason', sub.auditCommissionRejectionReason || '-'],
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
                <td>${label === 'Customer Account Number' ? renderCopyableCustomerAccount(sub) : escapeHtml(value)}</td>
            </tr>
        `).join('');
    }
    applicationDetailsModal?.classList.add('active');
}

function closeApplicationDetailsModal() {
    applicationDetailsModal?.classList.remove('active');
}

function getLatestAuditCorrectionDocument(submission = {}) {
    const docs = Array.isArray(submission.auditCommissionCorrectionDocuments)
        ? submission.auditCommissionCorrectionDocuments
        : [];
    return docs.length ? docs[docs.length - 1] : null;
}

function getAuditCorrectionSummaryHtml(submission = {}) {
    const latestDoc = getLatestAuditCorrectionDocument(submission);
    const comment = String(submission.auditCommissionResubmitComment || latestDoc?.comment || '').trim();
    if (!latestDoc && !comment) return '';

    const docLink = latestDoc?.fileUrl
        ? `<a href="${escapeHtml(latestDoc.fileUrl)}" target="_blank" rel="noopener">${escapeHtml(latestDoc.name || latestDoc.fileName || 'Correction document')}</a>`
        : escapeHtml(latestDoc?.name || latestDoc?.fileName || 'No document link');

    return `
        <div class="audit-correction-summary">
            <strong>Correction Submitted</strong>
            ${comment ? `<p>${escapeHtml(comment)}</p>` : ''}
            ${latestDoc ? `<div><i class="fas fa-paperclip"></i> ${docLink}</div>` : ''}
        </div>
    `;
}

function showAuditActionModal({ mode = 'accept', submission = {} } = {}) {
    return new Promise((resolve) => {
        const isReject = mode === 'reject';
        const modal = document.createElement('div');
        modal.className = 'modal active audit-action-modal';
        modal.innerHTML = `
            <div class="modal-content audit-action-card ${isReject ? 'reject' : 'accept'}">
                <div class="audit-action-icon">
                    <i class="fas ${isReject ? 'fa-xmark' : 'fa-check'}"></i>
                </div>
                <h2>${isReject ? 'Reject Payment Request' : 'Approve Payment Request'}</h2>
                <p>${isReject ? 'Enter a reason for rejecting' : 'Confirm that commission payment should be marked as paid for'} <strong>${escapeHtml(submission.customerName || 'this application')}</strong>.</p>
                ${getAuditCorrectionSummaryHtml(submission)}
                ${isReject ? '<textarea id="auditRejectReasonInput" rows="4" placeholder="Enter rejection reason"></textarea>' : ''}
                <div class="audit-action-actions">
                    <button type="button" class="cancel-btn" data-audit-action="cancel">Cancel</button>
                    <button type="button" class="submit-btn ${isReject ? 'danger' : ''}" data-audit-action="confirm">
                        <i class="fas ${isReject ? 'fa-paper-plane' : 'fa-check'}"></i> ${isReject ? 'Reject' : 'Approve'}
                    </button>
                </div>
            </div>
        `;
        const close = (value) => {
            modal.remove();
            resolve(value);
        };
        modal.addEventListener('click', (event) => {
            if (event.target === modal) {
                close(isReject ? '' : false);
                return;
            }
            const button = event.target.closest('[data-audit-action]');
            if (!button) return;
            if (button.dataset.auditAction === 'cancel') {
                close(isReject ? '' : false);
                return;
            }
            if (isReject) {
                const reason = String(modal.querySelector('#auditRejectReasonInput')?.value || '').trim();
                if (!reason) {
                    modal.querySelector('#auditRejectReasonInput')?.classList.add('invalid');
                    return;
                }
                close(reason);
            } else {
                close(true);
            }
        });
        document.body.appendChild(modal);
        if (isReject) setTimeout(() => modal.querySelector('#auditRejectReasonInput')?.focus(), 0);
    });
}

window.openMonitoringApplicationDetails = openApplicationDetailsModal;
window.acceptAuditCommission = async (submissionId) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const confirmed = await showAuditActionModal({ mode: 'accept', submission: sub });
    if (!confirmed) return;

    try {
        const { commission, twentyFive, rsaBalance } = getSubmissionFinancials(sub);
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentUser?.email || '',
            commissionPaid: true,
            commissionPaidAt: serverTimestamp(),
            commissionPaidBy: currentUser?.email || '',
            auditCommissionStatus: 'accepted',
            auditCommissionAcceptedAt: serverTimestamp(),
            auditCommissionAcceptedBy: currentUser?.email || '',
            auditCommissionAmount: commission,
            auditRsaBalance: rsaBalance,
            auditRsaTwentyFivePercent: twentyFive,
            updatedAt: serverTimestamp()
        });

        await addDoc(collection(db, 'audit'), {
            action: 'audit_commission_accepted',
            submissionId,
            customerName: sub.customerName || '',
            uploadedBy: sub.uploadedBy || '',
            commissionAmount: commission,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        }).catch(() => {});

        await notifyUserPushEvent({
            currentUser,
            recipientEmail: String(sub.auditCommissionSubmittedBy || sub.paymentMadeBy || sub.uploadedBy || '').trim(),
            eventType: 'audit_commission_accepted',
            title: 'Payment Request Approved',
            body: `${sub.customerName || 'Your application'} has been marked as paid by Audit.`,
            clickUrl: '/dashboard.html',
            meta: {
                submissionId,
                customerName: sub.customerName || '',
                approvedBy: currentUser?.email || '',
                commissionAmount: commission
            }
        }).catch(() => {});

        showNotification('Commission accepted and marked paid.', 'success');
    } catch (error) {
        showNotification('Failed to accept commission.', 'error');
    }
};

window.rejectAuditCommission = async (submissionId) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const reason = String(await showAuditActionModal({ mode: 'reject', submission: sub }) || '').trim();
    if (!reason) {
        showNotification('Rejection reason is required.', 'warning');
        return;
    }

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            paymentMadeByUploader: false,
            auditCommissionStatus: 'rejected',
            auditCommissionRejectionReason: reason,
            auditCommissionRejectedAt: serverTimestamp(),
            auditCommissionRejectedBy: currentUser?.email || '',
            auditCommissionRejections: arrayUnion({
                reason,
                rejectedBy: currentUser?.email || '',
                rejectedAt: new Date().toISOString()
            }),
            updatedAt: serverTimestamp()
        });

        await addDoc(collection(db, 'audit'), {
            action: 'audit_commission_rejected',
            submissionId,
            customerName: sub.customerName || '',
            uploadedBy: sub.uploadedBy || '',
            reason,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        }).catch(() => {});

        await notifyUserPushEvent({
            currentUser,
            recipientEmail: String(sub.auditCommissionSubmittedBy || sub.paymentMadeBy || sub.uploadedBy || '').trim(),
            eventType: 'audit_commission_rejected',
            title: 'Payment Request Rejected',
            body: `${sub.customerName || 'Your application'} was rejected by Audit: ${reason}`,
            clickUrl: '/dashboard.html',
            meta: {
                submissionId,
                customerName: sub.customerName || '',
                rejectedBy: currentUser?.email || '',
                reason
            }
        }).catch(() => {});

        showNotification('Commission rejected and returned to uploader.', 'success');
    } catch (error) {
        showNotification('Failed to reject commission.', 'error');
    }
};

window.signOutUser = async () => {
    await performAppLogout({
        auth,
        beforeSignOut: async () => {
            const userId = currentUserData?.id || currentUser?.uid || '';
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

function bindEvents() {
    document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            switchTab(item.dataset.tab || 'overview');
        });
    });
    document.querySelectorAll('[data-audit-paid-scope]').forEach((button) => {
        button.addEventListener('click', () => {
            currentAuditPaidScope = button.dataset.auditPaidScope === 'others' ? 'others' : 'mine';
            renderAuditWorkflowTabs();
        });
    });

    [usersSearch, usersRoleFilter, usersStatusFilter, applicationsSearch, applicationsStatusFilter, applicationsStageFilter]
        .forEach((el) => el?.addEventListener('input', renderCurrentTab));
    [usersRoleFilter, usersStatusFilter, applicationsStatusFilter, applicationsStageFilter]
        .forEach((el) => el?.addEventListener('change', renderCurrentTab));

    const handleSignOut = (event) => {
        event.preventDefault();
        if (typeof window.signOutUser === 'function') {
            window.signOutUser();
        }
    };
    document.getElementById('signOutBtnSidebar')?.addEventListener('click', handleSignOut);
    document.getElementById('signOutBtnMobile')?.addEventListener('click', handleSignOut);
    document.getElementById('forceRefreshBtn')?.addEventListener('click', () => window.location.reload());
    document.getElementById('forceRefreshBtnMobile')?.addEventListener('click', () => window.location.reload());
    exportAuditPendingReportBtn?.addEventListener('click', () => {
        pendingPaymentReportRequest = { kind: 'sent-to-pfa' };
        openPaymentReportRangeModal();
    });
    exportAuditPaidReportBtn?.addEventListener('click', () => {
        pendingPaymentReportRequest = { kind: 'paid' };
        openPaymentReportRangeModal();
    });
    exportAuditClearedReportBtn?.addEventListener('click', () => {
        pendingPaymentReportRequest = { kind: 'cleared' };
        openPaymentReportRangeModal();
    });
    exportAuditRejectedReportBtn?.addEventListener('click', () => {
        pendingPaymentReportRequest = { kind: 'rejected' };
        openPaymentReportRangeModal();
    });
    document.getElementById('closePaymentReportRangeModalBtn')?.addEventListener('click', closePaymentReportRangeModal);
    document.getElementById('cancelPaymentReportRangeBtn')?.addEventListener('click', closePaymentReportRangeModal);
    document.getElementById('closePaymentReportPreviewModalBtn')?.addEventListener('click', closePaymentReportPreviewModal);
    document.getElementById('closePaymentReportPreviewFooterBtn')?.addEventListener('click', closePaymentReportPreviewModal);
    document.getElementById('closePaymentPdfPreviewModalBtn')?.addEventListener('click', closePaymentPdfPreviewModal);
    document.getElementById('closePaymentPdfPreviewFooterBtn')?.addEventListener('click', closePaymentPdfPreviewModal);
    downloadPaymentPdfPreviewBtn?.addEventListener('click', () => {
        if (currentPaymentPdfPreviewBlob) saveBlob(currentPaymentPdfPreviewFileName || 'payment-report.pdf', currentPaymentPdfPreviewBlob);
    });
    document.querySelectorAll('.payment-report-preview-tab').forEach((btn) => {
        btn.addEventListener('click', () => showPaymentReportPreviewTab(btn.dataset.paymentReportView || 'summary'));
    });
    generatePaymentReportBtn?.addEventListener('click', async () => {
        const startDate = paymentReportStartDate?.value || '';
        const endDate = paymentReportEndDate?.value || '';
        if (!startDate || !endDate) {
            showNotification('Please choose both start and end date.', 'warning');
            return;
        }
        if (startDate > endDate) {
            showNotification('Start date cannot be after end date.', 'warning');
            return;
        }
        const originalHtml = generatePaymentReportBtn.innerHTML;
        generatePaymentReportBtn.disabled = true;
        generatePaymentReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
        try {
            const report = createStageReportFromDateRange(pendingPaymentReportRequest, startDate, endDate);
            renderPaymentReportPreview(report);
            closePaymentReportRangeModal();
            openPaymentReportPreviewModal();
        } catch (error) {
            showNotification('Failed to generate payment report', 'error');
        } finally {
            generatePaymentReportBtn.disabled = false;
            generatePaymentReportBtn.innerHTML = originalHtml;
        }
    });
    exportPaymentReportExcelBtn?.addEventListener('click', async () => {
        if (!currentPaymentReport) {
            showNotification('Generate a report first.', 'warning');
            return;
        }
        const originalHtml = exportPaymentReportExcelBtn.innerHTML;
        exportPaymentReportExcelBtn.disabled = true;
        exportPaymentReportExcelBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Exporting...';
        try {
            await exportPaymentReportExcel(currentPaymentReport);
        } catch (_) {
            showNotification('Failed to export Excel report', 'error');
        } finally {
            exportPaymentReportExcelBtn.disabled = false;
            exportPaymentReportExcelBtn.innerHTML = originalHtml;
        }
    });
    exportPaymentReportPdfBtn?.addEventListener('click', async () => {
        if (!currentPaymentReport) {
            showNotification('Generate a report first.', 'warning');
            return;
        }
        const originalHtml = exportPaymentReportPdfBtn.innerHTML;
        exportPaymentReportPdfBtn.disabled = true;
        exportPaymentReportPdfBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Exporting...';
        try {
            await exportPaymentReportPdf(currentPaymentReport);
        } catch (error) {
            showNotification(error?.message || 'Failed to export PDF report', 'error');
        } finally {
            exportPaymentReportPdfBtn.disabled = false;
            exportPaymentReportPdfBtn.innerHTML = originalHtml;
        }
    });

    document.addEventListener('click', async (event) => {
        const copyButton = event.target.closest('[data-copy-account]');
        if (!copyButton) return;
        const accountNumber = String(copyButton.dataset.copyAccount || '').trim();
        try {
            const copied = await copyTextToClipboard(accountNumber);
            if (!copied) throw new Error('Copy failed');
            const icon = copyButton.querySelector('i');
            if (icon) {
                icon.className = 'fas fa-check';
                setTimeout(() => { icon.className = 'fas fa-copy'; }, 1200);
            }
            showNotification('Account number copied.', 'success');
        } catch (_) {
            showNotification('Could not copy account number.', 'error');
        }
    });

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
