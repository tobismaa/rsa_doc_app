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
import {
    getSubmissionPaymentEntryAt,
    getSubmissionPaidEntryAt,
    getSubmissionClearedEntryAt
} from './shared/submission-stage.js?v=20260609a';
import {
    buildDashboardStageReport,
    renderDashboardStageReport,
    exportDashboardStageReportExcel
} from './shared/dashboard-stage-report.js?v=20260610a';

let currentUser = null;
let currentUserData = null;
let allSubmissions = [];
let paymentLeaveHistoryLoaded = false;
let paymentMyLeaveHistory = [];
let paymentReliefLeaveHistory = [];
let paymentReconciliationSourceRows = [];
let paymentReconciliationResult = null;
let paymentReconciliationFileName = '';
let paymentReconciliationSelectedIds = new Set();
const uploaderNameCache = new Map();
let activePaymentTab = 'dashboard';
let currentPaymentReport = null;
let currentPaymentStageReport = null;
let pendingPaymentReportRequest = { kind: 'range' };
const PAYMENT_RATE_CUTOFF_MS = new Date('2026-05-07T00:00:00+01:00').getTime();

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
const paymentReconciliationFileInput = document.getElementById('paymentReconciliationFileInput');
const openPaymentReconciliationModalBtn = document.getElementById('openPaymentReconciliationModalBtn');
const paymentReconciliationModal = document.getElementById('paymentReconciliationModal');
const closePaymentReconciliationModalBtn = document.getElementById('closePaymentReconciliationModalBtn');
const closePaymentReconciliationModalFooterBtn = document.getElementById('closePaymentReconciliationModalFooterBtn');
const paymentReconciliationTemplateBtn = document.getElementById('paymentReconciliationTemplateBtn');
const paymentReconciliationSelectBtn = document.getElementById('paymentReconciliationSelectBtn');
const paymentReconciliationRunBtn = document.getElementById('paymentReconciliationRunBtn');
const paymentReconciliationMarkMatchedBtn = document.getElementById('paymentReconciliationMarkMatchedBtn');
const paymentReconciliationClearBtn = document.getElementById('paymentReconciliationClearBtn');
const paymentReconciliationFileMeta = document.getElementById('paymentReconciliationFileMeta');
const paymentReconciliationSummary = document.getElementById('paymentReconciliationSummary');
const paymentReconciliationNotes = document.getElementById('paymentReconciliationNotes');
const paymentReconciliationMatchedWrap = document.getElementById('paymentReconciliationMatchedWrap');
const paymentReconciliationMatchedBody = document.getElementById('paymentReconciliationMatchedBody');
const paymentReconciliationSelectAll = document.getElementById('paymentReconciliationSelectAll');
const paymentReconciliationUnmatchedWrap = document.getElementById('paymentReconciliationUnmatchedWrap');
const paymentReconciliationUnmatchedBody = document.getElementById('paymentReconciliationUnmatchedBody');
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
const paymentQueueSortFilter = document.getElementById('paymentQueueSortFilter');
const paymentPaidSortFilter = document.getElementById('paymentPaidSortFilter');
const paymentClearedSortFilter = document.getElementById('paymentClearedSortFilter');
const generateSentToPfaReportBtn = document.getElementById('generateSentToPfaReportBtn');
const generatePaidReportBtn = document.getElementById('generatePaidReportBtn');
const generateClearedReportBtn = document.getElementById('generateClearedReportBtn');
const paymentStageReportMeta = document.getElementById('paymentStageReportMeta');
const paymentStageReportStartDate = document.getElementById('paymentStageReportStartDate');
const paymentStageReportEndDate = document.getElementById('paymentStageReportEndDate');
const paymentStageReportSummaryBody = document.getElementById('paymentStageReportSummaryBody');
const paymentStageReportDetailsBody = document.getElementById('paymentStageReportDetailsBody');
const generatePaymentStageReportBtn = document.getElementById('generatePaymentStageReportBtn');
const exportPaymentStageReportBtn = document.getElementById('exportPaymentStageReportBtn');

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

function getTimestampMillis(value) {
    if (!value) return 0;
    try {
        if (typeof value.toMillis === 'function') return value.toMillis();
        if (typeof value.toDate === 'function') return value.toDate().getTime();
        const parsed = new Date(value);
        return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
    } catch (_) {
        return 0;
    }
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

function getGenericDisplayName(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return 'Unassigned';
    return uploaderNameCache.get(normalized) || normalized.split('@')[0] || normalized;
}

function initializePaymentStageReportDates() {
    const today = new Date();
    const sixDaysAgo = new Date(today.getTime() - (6 * 24 * 60 * 60 * 1000));
    const formatInput = (date) => {
        const year = date.getFullYear();
        const month = `${date.getMonth() + 1}`.padStart(2, '0');
        const day = `${date.getDate()}`.padStart(2, '0');
        return `${year}-${month}-${day}`;
    };
    if (paymentStageReportStartDate && !paymentStageReportStartDate.value) paymentStageReportStartDate.value = formatInput(sixDaysAgo);
    if (paymentStageReportEndDate && !paymentStageReportEndDate.value) paymentStageReportEndDate.value = formatInput(today);
}

function buildPaymentStageDashboardReport() {
    const startDate = String(paymentStageReportStartDate?.value || '').trim();
    const endDate = String(paymentStageReportEndDate?.value || '').trim();
    if (!startDate || !endDate) {
        throw new Error('Choose both start date and end date.');
    }
    if (startDate > endDate) {
        throw new Error('Start date cannot be after end date.');
    }
    const sourceRecords = allSubmissions.filter((sub) => (
        String(sub.status || '').toLowerCase() !== 'draft'
        && normalizeEmail(sub.assignedToPayment)
    ));
    return buildDashboardStageReport({
        stageId: 'payment',
        records: sourceRecords,
        rangeStart: startDate,
        rangeEnd: endDate,
        resolveName: getGenericDisplayName
    });
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

function getCellText(value) {
    if (value === undefined || value === null) return '';
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
    if (typeof value === 'object') {
        if (value.text) return String(value.text).trim();
        if (value.result !== undefined && value.result !== null) return String(value.result).trim();
    }
    return String(value).trim();
}

function normalizeImportHeader(header) {
    return String(header || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function normalizeCustomerName(value) {
    return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}

function normalizeMatchAmount(value) {
    return roundDownToNearestThousand(parseMoney(value));
}

function buildReconciliationKey(name, amount) {
    return `${normalizeCustomerName(name)}::${normalizeMatchAmount(amount)}`;
}

function buildNameOnlyKey(name) {
    return normalizeCustomerName(name);
}

function getSelectedPaymentReconciliationFile() {
    return paymentReconciliationFileInput?.files?.[0] || null;
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

function getSubmissionUploadedAtMillis(sub) {
    return getTimestampMillis(sub?.uploadedAt || sub?.createdAt || sub?.updatedAt);
}

function getPaymentReceivedAtMillis(sub) {
    return getTimestampMillis(getSubmissionPaymentEntryAt(sub));
}

function getPaymentTableEntryMillis(sub) {
    const status = String(sub?.status || '').toLowerCase();
    if (status === 'paid') return getTimestampMillis(getSubmissionPaidEntryAt(sub));
    if (status === 'cleared') return getTimestampMillis(getSubmissionClearedEntryAt(sub));
    return getPaymentReceivedAtMillis(sub);
}

function getSubmissionRatePercent(sub) {
    const uploadedAtMs = getSubmissionUploadedAtMillis(sub);
    if (uploadedAtMs && uploadedAtMs < PAYMENT_RATE_CUTOFF_MS) return 10;
    return 7;
}

function getSubmissionRateLabel(sub) {
    return `${getSubmissionRatePercent(sub)}%`;
}

function isPaymentSubmissionAttended(sub) {
    const status = String(sub?.status || '').toLowerCase();
    return status === 'paid' || status === 'cleared';
}

function getPaymentStatusLabel(sub) {
    const status = String(sub?.status || '').toLowerCase();
    if (status === 'cleared') return 'Cleared';
    if (status === 'paid') return 'Paid';
    return 'Sent to PFA';
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
    const emails = [...new Set(records.flatMap((sub) => [
        normalizeEmail(sub?.uploadedBy),
        normalizeEmail(sub?.assignedToPayment),
        normalizeEmail(sub?.paidBy),
        normalizeEmail(sub?.clearedBy),
        normalizeEmail(sub?.assignedToRSA)
    ]).filter(Boolean))];
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
            latestPaidBy: '',
            latestClearedAt: null,
            latestClearedBy: '',
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
        if (Number.isFinite(paidAtMs) && paidAtMs > 0 && (!existing.latestPaidAt || paidAtMs > existing.latestPaidAt)) {
            existing.latestPaidAt = paidAtMs;
            existing.latestPaidBy = String(sub?.paidBy || '').trim();
        }
        if (Number.isFinite(clearedAtMs) && clearedAtMs > 0 && (!existing.latestClearedAt || clearedAtMs > existing.latestClearedAt)) {
            existing.latestClearedAt = clearedAtMs;
            existing.latestClearedBy = String(sub?.clearedBy || '').trim();
        }
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
        const reconciliationMatch = getPaymentReconciliationMatch(sub.id);
        const reconciliationBadge = reconciliationMatch
            ? `<div style="margin-top:6px;"><span style="display:inline-flex;align-items:center;gap:6px;padding:4px 8px;border-radius:999px;background:#dcfce7;color:#166534;font-size:11px;font-weight:700;"><i class="fas fa-link"></i> Excel matched</span></div>`
            : '';
        const dateLabel = mode === 'cleared'
            ? formatDateValue(sub?.clearedAt || sub?.updatedAt)
            : mode === 'paid'
                ? formatDateValue(sub?.paidAt || sub?.updatedAt)
                : formatDateValue(sub?.rsaSubmittedAt || sub?.updatedAt);

        return `
            <tr style="${reconciliationMatch ? 'background:rgba(220,252,231,0.28);' : ''}">
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong>${reconciliationBadge}</td>
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
    activePaymentTab = tabId;
    document.querySelectorAll('.nav-item').forEach((nav) => nav.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');
    const titles = {
        dashboard: 'Payment Dashboard',
        'sent-to-pfa': 'Sent to PFA',
        'paid-customers': 'Paid',
        'cleared-customers': 'Cleared',
        report: 'Payment Report',
        leave: 'Leave History',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Payment Dashboard';
    if (tabId === 'leave') {
        renderPaymentLeaveHistory().catch(() => {});
    }
    if (tabId === 'report') {
        initializePaymentStageReportDates();
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
            <td><button class="action-btn" onclick="window.openPaymentLeaveApplications('${record.id}', this)"><i class="fas fa-eye"></i> View</button></td>
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

window.openPaymentLeaveApplications = async (recordId, buttonEl = null) => {
    const record = [...paymentMyLeaveHistory, ...paymentReliefLeaveHistory].find((item) => item.id === recordId);
    if (!record) return;
    const originalHtml = buttonEl?.innerHTML || '';
    try {
        if (buttonEl) {
            buttonEl.disabled = true;
            buttonEl.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading...';
        }
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
    } catch (_) {
        showNotification('Failed to load leave applications', 'error');
    } finally {
        if (buttonEl) {
            buttonEl.disabled = false;
            buttonEl.innerHTML = originalHtml;
        }
    }
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
        if (status === 'sent_to_pfa' || status === 'rsa_submitted') return true;
        if (status === 'paid' || status === 'cleared') return false;
        return sub.finalSubmitted === true || sub.rsaSubmitted === true;
    });
}

function getPaidRecords() {
    return getPaymentRecords().filter((sub) => String(sub.status || '').toLowerCase() === 'paid');
}

function getClearedRecords() {
    return allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() === 'cleared');
}

function getPaymentReportWorkflowRecords() {
    return getPaymentRecords()
        .slice()
        .sort((a, b) => getPaymentReceivedAtMillis(a) - getPaymentReceivedAtMillis(b));
}

function getSortedFilteredSentToPfaRecords() {
    const mode = String(paymentQueueSortFilter?.value || 'all').trim();
    const queue = getSentToPfaRecords().slice();
    return sortAndFilterPaymentRecords(queue, mode);
}

function sortAndFilterPaymentRecords(records = [], mode = 'all') {
    const queue = records
        .slice()
        .sort((a, b) => getPaymentTableEntryMillis(b) - getPaymentTableEntryMillis(a));

    if (mode === 'rate_7') {
        return queue.filter((sub) => getSubmissionRatePercent(sub) === 7);
    }
    if (mode === 'rate_10') {
        return queue.filter((sub) => getSubmissionRatePercent(sub) === 10);
    }
    if (mode === 'agent') {
        return queue.sort((a, b) => {
            const aName = String(a?.agentName || 'No Agent').trim().toLowerCase();
            const bName = String(b?.agentName || 'No Agent').trim().toLowerCase();
            return aName.localeCompare(bName);
        });
    }
    if (mode === 'uploader') {
        return queue.sort((a, b) => {
            const aName = String(getUploaderDisplayName(a?.uploadedBy) || '-').trim().toLowerCase();
            const bName = String(getUploaderDisplayName(b?.uploadedBy) || '-').trim().toLowerCase();
            return aName.localeCompare(bName);
        });
    }

    return queue;
}

function getGroupRateSummaryLabel(group) {
    const uniqueRates = [...new Set((group?.submissions || []).map((sub) => getSubmissionRatePercent(sub)))].sort((a, b) => a - b);
    if (!uniqueRates.length) return 'Rate: -';
    if (uniqueRates.length === 1) return `Rate: ${uniqueRates[0]}%`;
    return `Rates: ${uniqueRates.map((rate) => `${rate}%`).join(', ')}`;
}

function getSortedFilteredPaidRecords() {
    const mode = String(paymentPaidSortFilter?.value || 'all').trim();
    return sortAndFilterPaymentRecords(getPaidRecords(), mode);
}

function getSortedFilteredClearedGroups() {
    const mode = String(paymentClearedSortFilter?.value || 'all').trim();
    let groups = buildAgentPaymentGroups(getClearedRecords());

    if (mode === 'rate_7') {
        groups = groups.filter((group) => group.submissions.some((sub) => getSubmissionRatePercent(sub) === 7));
    } else if (mode === 'rate_10') {
        groups = groups.filter((group) => group.submissions.some((sub) => getSubmissionRatePercent(sub) === 10));
    } else if (mode === 'agent') {
        groups = groups.sort((a, b) => String(a.agentName || '').trim().toLowerCase().localeCompare(String(b.agentName || '').trim().toLowerCase()));
    } else if (mode === 'uploader') {
        groups = groups.sort((a, b) => String(a.uploaderName || a.uploaderEmail || '').trim().toLowerCase().localeCompare(String(b.uploaderName || b.uploaderEmail || '').trim().toLowerCase()));
    }

    return groups;
}

function getRangeBoundaryMs(dateValue, mode = 'start') {
    if (!dateValue) return 0;
    const suffix = mode === 'end' ? 'T23:59:59.999+01:00' : 'T00:00:00+01:00';
    return new Date(`${dateValue}${suffix}`).getTime();
}

function buildPaymentReport(rangeStart, rangeEnd) {
    const startMs = getRangeBoundaryMs(rangeStart, 'start');
    const endMs = getRangeBoundaryMs(rangeEnd, 'end');
    const filteredRecords = getPaymentReportWorkflowRecords().filter((sub) => {
        const receivedAtMs = getPaymentReceivedAtMillis(sub);
        return receivedAtMs >= startMs && receivedAtMs <= endMs;
    });

    const summaryMap = new Map();
    const details = filteredRecords
        .map((sub) => {
            const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
            const receivedAtMs = getPaymentReceivedAtMillis(sub);
            const attended = isPaymentSubmissionAttended(sub);
            const statusLabel = getPaymentStatusLabel(sub);
            const detailRow = {
                id: sub.id,
                receivedAtMs,
                receivedDate: formatDateOnly(receivedAtMs),
                customerName: String(sub?.customerName || 'Unknown').trim() || 'Unknown',
                agentName: String(sub?.agentName || '').trim() || 'No Agent',
                uploaderName: getUploaderDisplayName(sub?.uploadedBy),
                pfa,
                twentyFive,
                commission: commission2,
                rateLabel: getSubmissionRateLabel(sub),
                statusLabel,
                attendedLabel: attended ? 'Yes' : 'No',
                paidDate: formatDateValue(sub?.paidAt),
                clearedDate: formatDateValue(sub?.clearedAt)
            };

            const summaryKey = formatInputDateValue(receivedAtMs);
            const existingSummary = summaryMap.get(summaryKey) || {
                key: summaryKey,
                dateLabel: formatDateOnly(receivedAtMs),
                received: 0,
                attended: 0,
                pending: 0
            };
            existingSummary.received += 1;
            if (attended) {
                existingSummary.attended += 1;
            } else {
                existingSummary.pending += 1;
            }
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
        title: 'Payment Report',
        exportKey: 'payment-report',
        rangeStart,
        rangeEnd,
        metaText: `Report range: ${formatDateOnly(getRangeBoundaryMs(rangeStart, 'start'))} to ${formatDateOnly(getRangeBoundaryMs(rangeEnd, 'end'))}.`,
        summaryRows,
        details,
        totals,
        dateCount: summaryRows.length
    };
}

function buildPaymentStageReport(records = [], options = {}) {
    const {
        title = 'Payment Report',
        exportKey = 'payment-report',
        dateLabel = 'Date',
        getDateMs = (sub) => getPaymentReceivedAtMillis(sub),
        statusResolver = (sub) => getPaymentStatusLabel(sub),
        attendedResolver = (sub) => isPaymentSubmissionAttended(sub)
    } = options;

    const summaryMap = new Map();
    const details = records
        .map((sub) => {
            const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
            const dateMs = getDateMs(sub);
            const attended = attendedResolver(sub);
            const statusLabel = statusResolver(sub);
            const detailRow = {
                id: sub.id,
                receivedAtMs: dateMs,
                receivedDate: formatDateOnly(dateMs),
                customerName: String(sub?.customerName || 'Unknown').trim() || 'Unknown',
                agentName: String(sub?.agentName || '').trim() || 'No Agent',
                uploaderName: getUploaderDisplayName(sub?.uploadedBy),
                pfa,
                twentyFive,
                commission: commission2,
                rateLabel: getSubmissionRateLabel(sub),
                statusLabel,
                attendedLabel: attended ? 'Yes' : 'No',
                paidDate: formatDateValue(sub?.paidAt),
                clearedDate: formatDateValue(sub?.clearedAt)
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
        metaText: `${title} generated from ${records.length} record(s). Summary grouped by ${dateLabel.toLowerCase()}.`,
        summaryRows,
        details,
        totals,
        dateCount: summaryRows.length
    };
}

function renderPaymentReportPreview(report) {
    currentPaymentReport = report;

    if (paymentReportPreviewMeta) {
        paymentReportPreviewMeta.textContent = report.metaText || 'Payment report preview.';
    }

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
            ? report.summaryRows.map((row) => `
                <tr>
                    <td>${escapeHtml(row.dateLabel)}</td>
                    <td>${escapeHtml(String(row.received))}</td>
                    <td>${escapeHtml(String(row.attended))}</td>
                    <td>${escapeHtml(String(row.pending))}</td>
                </tr>
            `).join('')
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
    summarySheet.addRow({
        dateLabel: 'Total',
        received: report.totals.received,
        attended: report.totals.attended,
        pending: report.totals.pending
    });
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
    const fileBase = report.exportKey || 'payment-report';
    saveBlob(`${fileBase}.xlsx`, new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
}

async function exportPaymentReportPdf(report) {
    if (!report || !window.jspdf?.jsPDF) {
        throw new Error('PDF library not available.');
    }

    const pdf = new window.jspdf.jsPDF({
        orientation: 'landscape',
        unit: 'pt',
        format: 'a4'
    });

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
        styles: {
            fontSize: 7,
            cellPadding: 2.5,
            overflow: 'linebreak',
            valign: 'middle'
        },
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

    pdf.save(`${report.exportKey || 'payment-report'}.pdf`);
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

function attachSortFilterToTableToolbar(selectId, tbodyId) {
    const select = document.getElementById(selectId);
    const tbody = document.getElementById(tbodyId);
    if (!select || !tbody) return;

    const table = tbody.closest('table');
    const container = table?.closest('.table-container');
    const controls = container?.previousElementSibling;
    if (!table || !container || !controls || !controls.classList.contains('table-enhancer-controls')) return;

    const wrapper = select.parentElement;
    const pager = controls.querySelector('.table-enhancer-pager');
    if (!wrapper || wrapper.parentElement === controls) return;

    wrapper.style.margin = '0';
    wrapper.style.display = 'flex';
    wrapper.style.alignItems = 'center';
    wrapper.style.gap = '10px';
    wrapper.style.flexWrap = 'wrap';

    if (pager) {
        controls.insertBefore(wrapper, pager);
    } else {
        controls.appendChild(wrapper);
    }
}

function syncPaymentTableToolbars() {
    window.requestAnimationFrame(() => {
        attachSortFilterToTableToolbar('paymentQueueSortFilter', 'paymentsTableBody');
        attachSortFilterToTableToolbar('paymentPaidSortFilter', 'paidCustomersTableBody');
        attachSortFilterToTableToolbar('paymentClearedSortFilter', 'clearedCustomersTableBody');
    });
}

function createStageReportFromDateRange(request, startDate, endDate) {
    const startMs = getRangeBoundaryMs(startDate, 'start');
    const endMs = getRangeBoundaryMs(endDate, 'end');

    if (request?.kind === 'sent-to-pfa') {
        const records = getSortedFilteredSentToPfaRecords().filter((sub) => {
            const dateMs = getPaymentReceivedAtMillis(sub);
            return dateMs >= startMs && dateMs <= endMs;
        });
        return buildPaymentStageReport(records, {
            title: 'Sent to PFA Report',
            exportKey: `sent-to-pfa-report-${startDate}-to-${endDate}`,
            dateLabel: 'sent date',
            getDateMs: (sub) => getPaymentReceivedAtMillis(sub),
            statusResolver: () => 'Sent to PFA',
            attendedResolver: () => false
        });
    }

    if (request?.kind === 'paid') {
        const records = getSortedFilteredPaidRecords().filter((sub) => {
            const dateMs = getTimestampMillis(getSubmissionPaidEntryAt(sub));
            return dateMs >= startMs && dateMs <= endMs;
        });
        return buildPaymentStageReport(records, {
            title: 'Paid Report',
            exportKey: `paid-report-${startDate}-to-${endDate}`,
            dateLabel: 'paid date',
            getDateMs: (sub) => getTimestampMillis(getSubmissionPaidEntryAt(sub)),
            statusResolver: () => 'Paid',
            attendedResolver: () => true
        });
    }

    if (request?.kind === 'cleared') {
        const records = getSortedFilteredClearedGroups()
            .flatMap((group) => group.submissions || [])
            .filter((sub) => {
                const dateMs = getTimestampMillis(getSubmissionClearedEntryAt(sub));
                return dateMs >= startMs && dateMs <= endMs;
            });
        return buildPaymentStageReport(records, {
            title: 'Cleared Report',
            exportKey: `cleared-report-${startDate}-to-${endDate}`,
            dateLabel: 'cleared date',
            getDateMs: (sub) => getTimestampMillis(getSubmissionClearedEntryAt(sub)),
            statusResolver: () => 'Cleared',
            attendedResolver: () => true
        });
    }

    return buildPaymentReport(startDate, endDate);
}

function resetPaymentReconciliationState() {
    paymentReconciliationSourceRows = [];
    paymentReconciliationResult = null;
    paymentReconciliationFileName = '';
    paymentReconciliationSelectedIds = new Set();
    if (paymentReconciliationFileInput) paymentReconciliationFileInput.value = '';
    renderPaymentReconciliation();
}

function getSelectableMatchedItems() {
    return paymentReconciliationResult?.matchedRows?.filter((item) => hasCommissionEligibleAgent(item.submission)) || [];
}

function syncPaymentReconciliationSelection() {
    const selectableIds = new Set(getSelectableMatchedItems().map((item) => item.submissionId));
    paymentReconciliationSelectedIds = new Set(
        [...paymentReconciliationSelectedIds].filter((id) => selectableIds.has(id))
    );
}

function updatePaymentReconciliationSelectAllState() {
    if (!paymentReconciliationSelectAll) return;
    const selectableItems = getSelectableMatchedItems();
    const selectableCount = selectableItems.length;
    const selectedCount = selectableItems.filter((item) => paymentReconciliationSelectedIds.has(item.submissionId)).length;
    paymentReconciliationSelectAll.checked = selectableCount > 0 && selectedCount === selectableCount;
    paymentReconciliationSelectAll.indeterminate = selectedCount > 0 && selectedCount < selectableCount;
    paymentReconciliationSelectAll.disabled = selectableCount === 0;
}

function togglePaymentReconciliationSelection(submissionId, checked) {
    if (!submissionId) return;
    if (checked) paymentReconciliationSelectedIds.add(submissionId);
    else paymentReconciliationSelectedIds.delete(submissionId);
    renderPaymentReconciliation();
}

function setPaymentReconciliationSelectionForAll(checked) {
    const selectableItems = getSelectableMatchedItems();
    paymentReconciliationSelectedIds = checked
        ? new Set(selectableItems.map((item) => item.submissionId))
        : new Set();
    renderPaymentReconciliation();
}

function openPaymentReconciliationModal() {
    paymentReconciliationModal?.classList.add('active');
}

function closePaymentReconciliationModal() {
    paymentReconciliationModal?.classList.remove('active');
}

async function downloadPaymentReconciliationTemplate() {
    try {
        if (!window.ExcelJS) throw new Error('Excel library is not available right now.');
        const workbook = new window.ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Payment Reconciliation');
        sheet.columns = [
            { header: 'Customer Name', key: 'customerName', width: 34 },
            { header: '25% Balance', key: 'balance', width: 18 }
        ];
        sheet.getRow(1).font = { bold: true };
        sheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'DCEBFA' }
        };
        sheet.addRow({ customerName: 'Sample Customer', balance: 250000 });
        sheet.addRow({ customerName: '', balance: '' });
        sheet.getCell('C1').value = 'Instruction';
        sheet.getCell('C2').value = 'Fill only Customer Name and 25% Balance columns.';
        sheet.getColumn(3).width = 55;

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'payment-reconciliation-template.xlsx';
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        showNotification('Template downloaded successfully.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to download template', 'error');
    }
}

function resolveReconciliationHeaderKey(headerMap, acceptedKeys = []) {
    for (const key of acceptedKeys) {
        for (const [colNumber, headerKey] of headerMap.entries()) {
            if (headerKey === key) return colNumber;
        }
    }
    return 0;
}

async function parsePaymentReconciliationFile(file) {
    if (!window.ExcelJS) throw new Error('Excel library is not available right now.');
    if (!file) throw new Error('Select an Excel file first.');
    const workbook = new window.ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error('No worksheet found in this Excel file.');

    const headerRow = sheet.getRow(1);
    const headerMap = new Map();
    headerRow.eachCell((cell, colNumber) => {
        const normalized = normalizeImportHeader(getCellText(cell.value));
        if (normalized) headerMap.set(colNumber, normalized);
    });

    const nameColumn = resolveReconciliationHeaderKey(headerMap, ['customername', 'customer', 'name', 'customerfullname']);
    const balanceColumn = resolveReconciliationHeaderKey(headerMap, ['25balance', '25percentbalance', 'rsa25percent', '25percent', 'twentyfivebalance', 'twentyfivepercent', 'balance25', '25rsabalance', 'rsa25balance']);
    if (!nameColumn || !balanceColumn) {
        throw new Error('Excel must contain customer name and 25% balance columns.');
    }

    const rows = [];
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        const customerName = getCellText(row.getCell(nameColumn).value);
        const rawBalance = getCellText(row.getCell(balanceColumn).value);
        if (!customerName && !rawBalance) continue;
        rows.push({
            rowNumber,
            customerName,
            rawBalance,
            normalizedName: normalizeCustomerName(customerName),
            normalizedAmount: normalizeMatchAmount(rawBalance)
        });
    }

    if (!rows.length) throw new Error('The Excel file does not contain any usable rows.');
    return rows;
}

function computePaymentReconciliationResult(rows = []) {
    const queue = getSentToPfaRecords();
    const queueBuckets = new Map();
    const nameBuckets = new Map();
    queue.forEach((sub) => {
        const { twentyFive } = getSubmissionFinancials(sub);
        const entry = {
            submission: sub,
            normalizedAmount: normalizeMatchAmount(twentyFive)
        };
        const key = buildReconciliationKey(sub?.customerName || '', twentyFive);
        const nameKey = buildNameOnlyKey(sub?.customerName || '');
        if (!queueBuckets.has(key)) queueBuckets.set(key, []);
        if (!nameBuckets.has(nameKey)) nameBuckets.set(nameKey, []);
        queueBuckets.get(key).push(entry);
        nameBuckets.get(nameKey).push(entry);
    });

    const matchedRows = [];
    const unmatchedRows = [];
    const matchedIds = new Set();
    const matchedBySubmissionId = new Map();

    rows.forEach((row) => {
        const exactKey = buildReconciliationKey(row.customerName, row.rawBalance);
        const exactMatches = queueBuckets.get(exactKey) || [];
        if (exactMatches.length) {
            const match = exactMatches.shift();
            matchedRows.push({
                ...row,
                submissionId: match.submission.id,
                submission: match.submission,
                queueAmount: match.normalizedAmount
            });
            matchedIds.add(match.submission.id);
            matchedBySubmissionId.set(match.submission.id, {
                rowNumber: row.rowNumber,
                customerName: row.customerName,
                normalizedAmount: row.normalizedAmount
            });
            return;
        }

        const nameMatches = nameBuckets.get(row.normalizedName) || [];
        unmatchedRows.push({
            ...row,
            reason: nameMatches.length
                ? 'Customer name exists, but the 25% balance does not match.'
                : 'Customer not found in the current Sent to PFA table.'
        });
    });

    return {
        totalRows: rows.length,
        matchedRows,
        unmatchedRows,
        matchedIds,
        matchedBySubmissionId,
        queueCount: queue.length
    };
}

function rerunPaymentReconciliation() {
    paymentReconciliationResult = paymentReconciliationSourceRows.length
        ? computePaymentReconciliationResult(paymentReconciliationSourceRows)
        : null;
    renderPaymentReconciliation();
}

function getPaymentReconciliationMatch(submissionId) {
    return paymentReconciliationResult?.matchedBySubmissionId?.get(submissionId) || null;
}

function renderPaymentReconciliation() {
    const selectedFile = getSelectedPaymentReconciliationFile();
    if (paymentReconciliationFileMeta) {
        paymentReconciliationFileMeta.textContent = paymentReconciliationFileName
            ? `Selected file: ${paymentReconciliationFileName}`
            : (selectedFile ? `Selected file: ${selectedFile.name}` : 'No Excel file selected.');
    }
    if (paymentReconciliationRunBtn) paymentReconciliationRunBtn.disabled = !selectedFile;
    if (paymentReconciliationClearBtn) paymentReconciliationClearBtn.disabled = !selectedFile && !paymentReconciliationResult;

    if (!paymentReconciliationResult) {
        if (paymentReconciliationSummary) {
            paymentReconciliationSummary.style.display = 'none';
            paymentReconciliationSummary.innerHTML = '';
        }
        if (paymentReconciliationNotes) {
            paymentReconciliationNotes.style.display = 'none';
            paymentReconciliationNotes.textContent = '';
        }
        if (paymentReconciliationMatchedWrap) paymentReconciliationMatchedWrap.style.display = 'none';
        if (paymentReconciliationMatchedBody) paymentReconciliationMatchedBody.innerHTML = '';
        if (paymentReconciliationUnmatchedWrap) paymentReconciliationUnmatchedWrap.style.display = 'none';
        if (paymentReconciliationUnmatchedBody) paymentReconciliationUnmatchedBody.innerHTML = '';
        if (paymentReconciliationMarkMatchedBtn) paymentReconciliationMarkMatchedBtn.disabled = true;
        return;
    }

    const matchedCount = paymentReconciliationResult.matchedRows.length;
    const unmatchedCount = paymentReconciliationResult.unmatchedRows.length;
    const actionableCount = paymentReconciliationResult.matchedRows.filter((item) => hasCommissionEligibleAgent(item.submission)).length;
    syncPaymentReconciliationSelection();
    const selectedCount = paymentReconciliationResult.matchedRows.filter((item) => paymentReconciliationSelectedIds.has(item.submissionId)).length;

    if (paymentReconciliationSummary) {
        paymentReconciliationSummary.style.display = 'flex';
        paymentReconciliationSummary.innerHTML = [
            { label: 'Excel Rows', value: paymentReconciliationResult.totalRows, bg: '#e2e8f0', color: '#334155' },
            { label: 'Matched', value: matchedCount, bg: '#dcfce7', color: '#166534' },
            { label: 'Not Found', value: unmatchedCount, bg: '#fee2e2', color: '#b91c1c' },
            { label: 'Queue Records', value: paymentReconciliationResult.queueCount, bg: '#dbeafe', color: '#1d4ed8' }
        ].map((item) => `<span style="display:inline-flex;align-items:center;gap:6px;padding:8px 12px;border-radius:999px;background:${item.bg};color:${item.color};font-size:12px;font-weight:700;"><span>${escapeHtml(String(item.label))}</span><span>${escapeHtml(String(item.value))}</span></span>`).join('');
    }

    if (paymentReconciliationNotes) {
        paymentReconciliationNotes.style.display = 'block';
        paymentReconciliationNotes.textContent = matchedCount
            ? `${matchedCount} row(s) matched automatically. Tick the applications you want to bulk mark as paid, or use single-row action for one-off handling.`
            : 'No queue rows matched this Excel file.';
    }

    if (paymentReconciliationMarkMatchedBtn) {
        paymentReconciliationMarkMatchedBtn.disabled = selectedCount === 0;
        paymentReconciliationMarkMatchedBtn.innerHTML = `<i class="fas fa-check-circle"></i> Mark Selected Paid${selectedCount ? ` (${selectedCount})` : ''}`;
    }
    updatePaymentReconciliationSelectAllState();
    if (paymentReconciliationMatchedWrap) paymentReconciliationMatchedWrap.style.display = matchedCount ? 'block' : 'none';
    if (paymentReconciliationMatchedBody) {
        paymentReconciliationMatchedBody.innerHTML = matchedCount
            ? paymentReconciliationResult.matchedRows.map((item) => {
                const sub = item.submission || {};
                const { pfa, twentyFive } = getSubmissionFinancials(sub);
                const canMarkPaid = hasCommissionEligibleAgent(sub);
                const checked = paymentReconciliationSelectedIds.has(sub.id) ? 'checked' : '';
                const checkboxHtml = canMarkPaid
                    ? `<input type="checkbox" class="payment-reconciliation-row-check" data-submission-id="${escapeHtml(sub.id)}" ${checked}>`
                    : `<span style="color:#94a3b8;font-size:12px;">N/A</span>`;
                const actionHtml = canMarkPaid
                    ? `<button class="action-btn" style="background:#16a34a;color:#fff;border:none;" onclick="window.markMatchedPaymentReconciliationRecord('${sub.id}')"><i class="fas fa-check-circle"></i> Mark Paid</button>`
                    : `<button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.markMatchedPaymentReconciliationRecord('${sub.id}')"><i class="fas fa-check-double"></i> Clear</button>`;
                return `
                    <tr>
                        <td>${checkboxHtml}</td>
                        <td>${escapeHtml(String(item.rowNumber))}</td>
                        <td><strong>${escapeHtml(sub.customerName || item.customerName || '-')}</strong></td>
                        <td>${escapeHtml(String(sub.agentName || '').trim() || 'No Agent')}</td>
                        <td>${escapeHtml(pfa)}</td>
                        <td>${formatCurrency(item.normalizedAmount || 0)}</td>
                        <td>${formatCurrency(twentyFive || 0)}</td>
                        <td><span class="status-badge status-approved">Matched</span></td>
                        <td>${actionHtml}</td>
                    </tr>
                `;
            }).join('')
            : '';
    }
    if (paymentReconciliationUnmatchedWrap) paymentReconciliationUnmatchedWrap.style.display = unmatchedCount ? 'block' : 'none';
    if (paymentReconciliationUnmatchedBody) {
        paymentReconciliationUnmatchedBody.innerHTML = unmatchedCount
            ? paymentReconciliationResult.unmatchedRows.map((row) => `
                <tr>
                    <td>${escapeHtml(String(row.rowNumber))}</td>
                    <td>${escapeHtml(row.customerName || '-')}</td>
                    <td>${escapeHtml(formatCurrency(row.normalizedAmount || 0))}</td>
                    <td>${escapeHtml(row.reason || 'Not found')}</td>
                </tr>
            `).join('')
            : '';
    }
}

function renderPaymentQueue() {
    if (!paymentsTableBody) return;

    const paymentQueue = getSortedFilteredSentToPfaRecords();

    if (paymentPendingCount) {
        paymentPendingCount.textContent = String(paymentQueue.length);
        paymentPendingCount.style.display = 'inline-block';
    }

    if (paymentQueue.length === 0) {
        paymentsTableBody.innerHTML = '<tr><td colspan="10" class="no-data">No applications sent to PFA yet</td></tr>';
        return;
    }

    paymentsTableBody.innerHTML = paymentQueue.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const queueDate = formatDateValue(getSubmissionPaymentEntryAt(sub));
        const uploaderLabel = getUploaderDisplayName(sub?.uploadedBy);
        const assignedLabel = getUploaderDisplayName(sub?.assignedToPayment);
        const agentName = String(sub?.agentName || '').trim() || 'No Agent';
        const rateLabel = getSubmissionRateLabel(sub);
        const reconciliationMatch = getPaymentReconciliationMatch(sub.id);
        const reconciliationBadge = reconciliationMatch
            ? `<div style="margin-top:6px;"><span style="display:inline-flex;align-items:center;gap:6px;padding:4px 8px;border-radius:999px;background:#dcfce7;color:#166534;font-size:11px;font-weight:700;"><i class="fas fa-link"></i> Excel matched</span></div>`
            : '';
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
                <td>${formatCurrency(twentyFive)}<div style="font-size:12px;color:#64748b;margin-top:4px;">Rate: ${escapeHtml(rateLabel)}</div></td>
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

function renderPaidCustomersSimpleTableLegacy() {
    if (!paidCustomersTableBody) return;

    const paidCustomers = getSortedFilteredPaidRecords();
    if (paidCustomerCount) {
        paidCustomerCount.textContent = String(paidCustomers.length);
        paidCustomerCount.style.display = 'inline-block';
    }

    if (!paidCustomers.length) {
        paidCustomersTableBody.innerHTML = '<tr><td colspan="10" class="no-data">No paid applications yet</td></tr>';
        return;
    }

    paidCustomersTableBody.innerHTML = paidCustomers.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const paidDate = formatDateValue(getSubmissionPaidEntryAt(sub));
        const uploaderLabel = getUploaderDisplayName(sub?.uploadedBy);
        const approvedByLabel = getUploaderDisplayName(sub?.paidBy);
        const agentName = String(sub?.agentName || '').trim() || 'No Agent';
        const rateLabel = getSubmissionRateLabel(sub);
        const clearAction = hasCommissionEligibleAgent(sub)
            ? `<button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearPaidAgent('${getAgentPaymentKey(sub)}')"><i class="fas fa-check-double"></i> Clear</button>`
            : `<button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearSubmissionWithoutAgent('${sub.id}')"><i class="fas fa-check-double"></i> Clear</button>`;

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>
                    <strong>${escapeHtml(agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(String(sub?.agentAccountBank || '').trim() || '-')} â€¢ ${escapeHtml(String(sub?.agentAccountNumber || '').trim() || '-')}</div>
                </td>
                <td>${escapeHtml(uploaderLabel || '-')}</td>
                <td>${escapeHtml(approvedByLabel || '-')}</td>
                <td>${escapeHtml(pfa)}</td>
                <td>${formatCurrency(twentyFive)}<div style="font-size:12px;color:#64748b;margin-top:4px;">Rate: ${escapeHtml(rateLabel)}</div></td>
                <td>${formatCurrency(commission2)}</td>
                <td>${escapeHtml(paidDate)}</td>
                <td><span class="status-badge status-approved">Paid</span></td>
                <td>${clearAction} <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');
}

// Override the paid-table renderer with an ASCII-only separator so account info
// displays reliably even if an earlier edit introduced mojibake.
function renderPaidCustomersSimpleTable() {
    if (!paidCustomersTableBody) return;

    const paidCustomers = getSortedFilteredPaidRecords();
    if (paidCustomerCount) {
        paidCustomerCount.textContent = String(paidCustomers.length);
        paidCustomerCount.style.display = 'inline-block';
    }

    if (!paidCustomers.length) {
        paidCustomersTableBody.innerHTML = '<tr><td colspan="10" class="no-data">No paid applications yet</td></tr>';
        return;
    }

    paidCustomersTableBody.innerHTML = paidCustomers.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const paidDate = formatDateValue(getSubmissionPaidEntryAt(sub));
        const uploaderLabel = getUploaderDisplayName(sub?.uploadedBy);
        const approvedByLabel = getUploaderDisplayName(sub?.paidBy);
        const agentName = String(sub?.agentName || '').trim() || 'No Agent';
        const rateLabel = getSubmissionRateLabel(sub);
        const clearAction = hasCommissionEligibleAgent(sub)
            ? `<button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearPaidAgent('${getAgentPaymentKey(sub)}')"><i class="fas fa-check-double"></i> Clear</button>`
            : `<button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearSubmissionWithoutAgent('${sub.id}')"><i class="fas fa-check-double"></i> Clear</button>`;

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>
                    <strong>${escapeHtml(agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(String(sub?.agentAccountBank || '').trim() || '-')} - ${escapeHtml(String(sub?.agentAccountNumber || '').trim() || '-')}</div>
                </td>
                <td>${escapeHtml(uploaderLabel || '-')}</td>
                <td>${escapeHtml(approvedByLabel || '-')}</td>
                <td>${escapeHtml(pfa)}</td>
                <td>${formatCurrency(twentyFive)}<div style="font-size:12px;color:#64748b;margin-top:4px;">Rate: ${escapeHtml(rateLabel)}</div></td>
                <td>${formatCurrency(commission2)}</td>
                <td>${escapeHtml(paidDate)}</td>
                <td><span class="status-badge status-approved">Paid</span></td>
                <td>${clearAction} <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');
}

function renderClearedCustomers() {
    if (!clearedCustomersTableBody) return;

    const groupedCleared = getSortedFilteredClearedGroups();
    if (clearedCustomerCount) {
        clearedCustomerCount.textContent = String(groupedCleared.length);
        clearedCustomerCount.style.display = 'inline-block';
    }

    if (groupedCleared.length === 0) {
        clearedCustomersTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No cleared agent batches yet</td></tr>';
        return;
    }

    clearedCustomersTableBody.innerHTML = groupedCleared.map((group) => {
        const breakdownId = `cleared-breakdown-${toSafeDomId(group.key)}`;
        const clearedDate = formatDateValue(group.latestClearedAt ? new Date(group.latestClearedAt) : null);
        const clearedByLabel = getUploaderDisplayName(group.latestClearedBy);
        const rateLabel = getGroupRateSummaryLabel(group);
        return `
            <tr>
                <td>
                    <strong>${escapeHtml(group.agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</div>
                </td>
                <td>${escapeHtml(group.uploaderName || group.uploaderEmail || '-')}</td>
                <td>${escapeHtml(clearedByLabel || '-')}</td>
                <td>${group.customerCount}</td>
                <td>${formatCurrency(group.total25)}<div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(rateLabel)}</div></td>
                <td>${formatCurrency(group.totalCommission)}</td>
                <td><span class="status-badge status-approved">Cleared</span><div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(clearedDate)}</div></td>
                <td><button class="action-btn agent-breakdown-toggle" onclick="window.togglePaymentAgentBreakdown('${breakdownId}', this)"><i class="fas fa-chevron-down"></i> Breakdown</button></td>
            </tr>
            <tr id="${breakdownId}" class="agent-breakdown-row" style="display:none;">
                <td colspan="8">${renderAgentBreakdownTable(group, 'paid')}</td>
            </tr>
        `;
    }).join('');
}

function loadSubmissions() {
    if (!currentUser) return;

    const q = query(collection(db, 'submissions'));
    onSnapshot(q, async (snapshot) => {
        allSubmissions = snapshot.docs
            .map((d) => ({ id: d.id, ...d.data() }))
            .sort((a, b) => {
                const aMs = getPaymentTableEntryMillis(a);
                const bMs = getPaymentTableEntryMillis(b);
                return bMs - aMs;
            });
        await primeUploaderNames(allSubmissions);
        renderDashboardOverview();
        rerunPaymentReconciliation();
        renderPaymentQueue();
        renderPaidCustomersSimpleTable();
        renderClearedCustomers();
        syncPaymentTableToolbars();
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

window.markMatchedPaymentReconciliationRecords = async () => {
    if (!paymentReconciliationResult?.matchedRows?.length) {
        showNotification('Run reconciliation first.', 'warning');
        return;
    }

    const matchedItems = paymentReconciliationResult.matchedRows.filter((item) => (
        hasCommissionEligibleAgent(item.submission) && paymentReconciliationSelectedIds.has(item.submissionId)
    ));
    if (!matchedItems.length) {
        showNotification('Tick at least one matched application to continue.', 'warning');
        return;
    }

    const confirmed = confirm(`Mark ${matchedItems.length} selected application(s) as paid?`);
    if (!confirmed) return;

    try {
        await Promise.all(matchedItems.map((item) => updateDoc(doc(db, 'submissions', item.submissionId), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentUser?.email || '',
            paymentReconciliationFileName: paymentReconciliationFileName || '',
            paymentReconciledAt: serverTimestamp()
        })));

        await addDoc(collection(db, 'audit'), {
            action: 'payment_reconciliation_bulk_paid',
            count: matchedItems.length,
            fileName: paymentReconciliationFileName || '',
            submissionIds: matchedItems.map((item) => item.submissionId),
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        paymentReconciliationSelectedIds = new Set(
            [...paymentReconciliationSelectedIds].filter((id) => !matchedItems.some((item) => item.submissionId === id))
        );
        showNotification(`Marked ${matchedItems.length} matched record(s) as paid`, 'success');
    } catch (error) {
        showNotification('Failed to mark matched records as paid', 'error');
    }
};

window.markMatchedPaymentReconciliationRecord = async (submissionId) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Matched application no longer exists in your queue.', 'warning');
        return;
    }
    if (hasCommissionEligibleAgent(sub)) {
        await window.markSubmissionPaid(submissionId);
        return;
    }
    await window.clearSubmissionWithoutAgent(submissionId);
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
    } catch (e) {}
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
openPaymentReconciliationModalBtn?.addEventListener('click', () => {
    openPaymentReconciliationModal();
});
closePaymentReconciliationModalBtn?.addEventListener('click', () => {
    closePaymentReconciliationModal();
});
closePaymentReconciliationModalFooterBtn?.addEventListener('click', () => {
    closePaymentReconciliationModal();
});
paymentReconciliationTemplateBtn?.addEventListener('click', () => {
    downloadPaymentReconciliationTemplate();
});
paymentReconciliationSelectBtn?.addEventListener('click', () => {
    paymentReconciliationFileInput?.click();
});
paymentReconciliationFileInput?.addEventListener('change', () => {
    paymentReconciliationFileName = getSelectedPaymentReconciliationFile()?.name || '';
    paymentReconciliationSourceRows = [];
    paymentReconciliationResult = null;
    paymentReconciliationSelectedIds = new Set();
    renderPaymentReconciliation();
    renderPaymentQueue();
});
paymentReconciliationRunBtn?.addEventListener('click', async () => {
    const file = getSelectedPaymentReconciliationFile();
    if (!file) {
        showNotification('Select an Excel file first.', 'warning');
        return;
    }
    try {
        paymentReconciliationRunBtn.disabled = true;
        paymentReconciliationRunBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Matching...';
        paymentReconciliationFileName = file.name;
        paymentReconciliationSourceRows = await parsePaymentReconciliationFile(file);
        rerunPaymentReconciliation();
        renderPaymentQueue();
        showNotification(`Reconciliation complete. ${paymentReconciliationResult?.matchedRows?.length || 0} row(s) matched.`, 'success');
    } catch (error) {
        paymentReconciliationSourceRows = [];
        paymentReconciliationResult = null;
        renderPaymentReconciliation();
        renderPaymentQueue();
        showNotification(error?.message || 'Failed to process Excel file', 'error');
    } finally {
        paymentReconciliationRunBtn.innerHTML = '<i class="fas fa-play"></i> Run Match';
        renderPaymentReconciliation();
    }
});
paymentReconciliationMarkMatchedBtn?.addEventListener('click', () => {
    window.markMatchedPaymentReconciliationRecords();
});
paymentReconciliationSelectAll?.addEventListener('change', (e) => {
    setPaymentReconciliationSelectionForAll(Boolean(e.target?.checked));
});
paymentReconciliationClearBtn?.addEventListener('click', () => {
    resetPaymentReconciliationState();
    renderPaymentQueue();
    showNotification('Reconciliation result cleared.', 'info');
});
paymentReconciliationMatchedBody?.addEventListener('change', (e) => {
    const target = e.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!target.classList.contains('payment-reconciliation-row-check')) return;
    togglePaymentReconciliationSelection(target.dataset.submissionId || '', target.checked);
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
document.getElementById('closePaymentReportRangeModalBtn')?.addEventListener('click', closePaymentReportRangeModal);
document.getElementById('cancelPaymentReportRangeBtn')?.addEventListener('click', closePaymentReportRangeModal);
document.getElementById('closePaymentReportPreviewModalBtn')?.addEventListener('click', closePaymentReportPreviewModal);
document.getElementById('closePaymentReportPreviewFooterBtn')?.addEventListener('click', closePaymentReportPreviewModal);
document.querySelectorAll('.payment-report-preview-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
        showPaymentReportPreviewTab(btn.dataset.paymentReportView || 'summary');
    });
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
generateSentToPfaReportBtn?.addEventListener('click', () => {
    pendingPaymentReportRequest = { kind: 'sent-to-pfa' };
    openPaymentReportRangeModal();
});
generatePaidReportBtn?.addEventListener('click', () => {
    pendingPaymentReportRequest = { kind: 'paid' };
    openPaymentReportRangeModal();
});
generateClearedReportBtn?.addEventListener('click', () => {
    pendingPaymentReportRequest = { kind: 'cleared' };
    openPaymentReportRangeModal();
});
generatePaymentStageReportBtn?.addEventListener('click', async () => {
    const originalHtml = generatePaymentStageReportBtn.innerHTML;
    generatePaymentStageReportBtn.disabled = true;
    generatePaymentStageReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    try {
        currentPaymentStageReport = buildPaymentStageDashboardReport();
        renderDashboardStageReport(currentPaymentStageReport, {
            metaEl: paymentStageReportMeta,
            summaryBodyEl: paymentStageReportSummaryBody,
            detailsBodyEl: paymentStageReportDetailsBody
        });
        showNotification('Payment report generated successfully.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to generate payment report.', 'error');
    } finally {
        generatePaymentStageReportBtn.disabled = false;
        generatePaymentStageReportBtn.innerHTML = originalHtml;
    }
});
exportPaymentStageReportBtn?.addEventListener('click', async () => {
    if (!currentPaymentStageReport) {
        showNotification('Generate a report first.', 'warning');
        return;
    }
    const originalHtml = exportPaymentStageReportBtn.innerHTML;
    exportPaymentStageReportBtn.disabled = true;
    exportPaymentStageReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Exporting...';
    try {
        await exportDashboardStageReportExcel(currentPaymentStageReport, 'CMBank RSA Payment Dashboard');
        showNotification('Payment report Excel downloaded.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to export payment report.', 'error');
    } finally {
        exportPaymentStageReportBtn.disabled = false;
        exportPaymentStageReportBtn.innerHTML = originalHtml;
    }
});
paymentQueueSortFilter?.addEventListener('change', () => {
    renderPaymentQueue();
});
paymentPaidSortFilter?.addEventListener('change', () => {
    renderPaidCustomersSimpleTable();
});
paymentClearedSortFilter?.addEventListener('change', () => {
    renderClearedCustomers();
});
window.addEventListener('click', (e) => {
    if (e.target === document.getElementById('paymentLeaveApplicationsModal')) {
        document.getElementById('paymentLeaveApplicationsModal')?.classList.remove('active');
    }
    if (e.target === paymentReconciliationModal) {
        closePaymentReconciliationModal();
    }
    if (e.target === paymentReportRangeModal) {
        closePaymentReportRangeModal();
    }
    if (e.target === paymentReportPreviewModal) {
        closePaymentReportPreviewModal();
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
        renderPaymentReconciliation();
        loadSubmissions();
        syncPaymentTableToolbars();
    } catch (error) {
        showNotification('Could not validate session', 'error');
        window.location.href = 'index.html';
    }
});
