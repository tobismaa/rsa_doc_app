import { auth, db } from './firebase-config.js?v=20260625c';
import { performAppLogout } from './shared/logout.js?v=20260625b';
import {
    collection,
    query,
    where,
    addDoc,
    doc,
    getDoc,
    onSnapshot,
    serverTimestamp,
    arrayUnion,
    updateDoc,
    writeBatch
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { formatAppDateTime } from './shared/app-time.js';
import { getCurrentUserProfile as getCurrentUserProfileShared } from './shared/user-directory.js?v=20260518a';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionOriginalUploadAt,
    getSubmissionCurrentStageEntryAt
} from './shared/submission-stage.js?v=20260714a';
import {
    getSubmissionCommissionAmount,
    resolveSubmissionCommissionRate
} from './shared/commission-config.js?v=20260507a';
import { notifyUserPushEvent } from './push-alerts.js';

let currentUser = null;
let currentUserData = null;
let allUsers = [];
let allSubmissions = [];
let userDisplayNamesByEmail = new Map();
let auditRenderTimer = null;
let currentTab = 'overview';
const AUDIT_DASHBOARD_TABS = ['overview', 'sent-to-pfa', 'paid', 'cleared', 'rejected', 'reconciliation', 'user-report', 'profile', 'help'];
const AUDIT_BULK_CLEAR_BATCH_SIZE = 200;
const AUDIT_RECONCILIATION_VIEWS = ['excel', 'duplicates', 'ignored', 'rejected'];

function getInitialAuditTab() {
    const hashTab = decodeURIComponent(String(window.location.hash || '').replace(/^#/, '')).trim();
    return AUDIT_DASHBOARD_TABS.includes(hashTab) ? hashTab : 'overview';
}

function rememberAuditTab(tabId) {
    if (!AUDIT_DASHBOARD_TABS.includes(tabId)) return;
    if (window.location.hash === `#${tabId}`) return;
    history.replaceState(null, '', `#${tabId}`);
}

function forceHardRefresh() {
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
}
let currentAuditPaidScope = 'all';
let currentPaymentReport = null;
let pendingPaymentReportRequest = { kind: 'paid' };
let currentPaymentPdfPreviewUrl = '';
let currentPaymentPdfPreviewBlob = null;
let currentPaymentPdfPreviewFileName = 'payment-report.pdf';
let auditReconciliationSourceRows = [];
let auditReconciliationResult = null;
let auditReconciliationFileName = '';
let auditDuplicateScanResult = null;
const auditDuplicateRestoreInFlight = new Set();
let auditReconciliationActiveView = 'excel';
let auditPaidReconciliationSourceRows = [];
let auditPaidReconciliationResult = null;
let auditPaidReconciliationFileName = '';
let auditPaidReconciliationSelectedSubmissionIds = [];
let auditPaidReconciliationActiveResultsTab = 'matched';
let currentAuditRejectedScope = 'rejected';
let currentAuditDuplicateHistoryFilter = 'all';
let usersListenerStarted = false;
let submissionsListenerMode = '';
let submissionListenerUnsubs = [];
const submissionSnapshotSources = new Map();
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
const auditReconciliationFileInput = document.getElementById('auditReconciliationFileInput');
const auditReconciliationTemplateBtn = document.getElementById('auditReconciliationTemplateBtn');
const auditReconciliationSelectBtn = document.getElementById('auditReconciliationSelectBtn');
const auditReconciliationRunBtn = document.getElementById('auditReconciliationRunBtn');
const auditReconciliationExportBtn = document.getElementById('auditReconciliationExportBtn');
const auditReconciliationClearBtn = document.getElementById('auditReconciliationClearBtn');
const auditReconciliationFileMeta = document.getElementById('auditReconciliationFileMeta');
const auditReconciliationSummary = document.getElementById('auditReconciliationSummary');
const auditDuplicateScanSummary = document.getElementById('auditDuplicateScanSummary');
const auditDuplicateScanWrap = document.getElementById('auditDuplicateScanWrap');
const auditDuplicateScanBody = document.getElementById('auditDuplicateScanBody');
const auditDuplicateIgnoredWrap = document.getElementById('auditDuplicateIgnoredWrap');
const auditDuplicateIgnoredBody = document.getElementById('auditDuplicateIgnoredBody');
const auditDuplicateRejectedWrap = document.getElementById('auditDuplicateRejectedWrap');
const auditDuplicateRejectedBody = document.getElementById('auditDuplicateRejectedBody');
const auditDuplicateHistoryFilter = document.getElementById('auditDuplicateHistoryFilter');
const auditUserReportTableBody = document.getElementById('auditUserReportTableBody');
const auditUserReportUserCount = document.getElementById('auditUserReportUserCount');
const auditUserReportSentCount = document.getElementById('auditUserReportSentCount');
const auditUserReportClearedCount = document.getElementById('auditUserReportClearedCount');
const auditUserAgentSummaryModal = document.getElementById('auditUserAgentSummaryModal');
const auditUserAgentSummaryTitle = document.getElementById('auditUserAgentSummaryTitle');
const auditUserAgentSummaryName = document.getElementById('auditUserAgentSummaryName');
const auditUserAgentSummaryMeta = document.getElementById('auditUserAgentSummaryMeta');
const auditUserAgentSummaryCards = document.getElementById('auditUserAgentSummaryCards');
const auditUserAgentSummaryBody = document.getElementById('auditUserAgentSummaryBody');
const auditReconciliationMatchedWrap = document.getElementById('auditReconciliationMatchedWrap');
const auditReconciliationMatchedBody = document.getElementById('auditReconciliationMatchedBody');
const auditReconciliationUnmatchedWrap = document.getElementById('auditReconciliationUnmatchedWrap');
const auditReconciliationUnmatchedBody = document.getElementById('auditReconciliationUnmatchedBody');
const openAuditPaidReconciliationBtn = document.getElementById('openAuditPaidReconciliationBtn');
const auditPaidReconciliationUploadModal = document.getElementById('auditPaidReconciliationUploadModal');
const auditPaidReconciliationResultsModal = document.getElementById('auditPaidReconciliationResultsModal');
const auditPaidReconciliationFileInput = document.getElementById('auditPaidReconciliationFileInput');
const auditPaidReconciliationSelectBtn = document.getElementById('auditPaidReconciliationSelectBtn');
const auditPaidReconciliationRunBtn = document.getElementById('auditPaidReconciliationRunBtn');
const auditPaidReconciliationSelectAllBtn = document.getElementById('auditPaidReconciliationSelectAllBtn');
const auditPaidReconciliationClearSelectedBtn = document.getElementById('auditPaidReconciliationClearSelectedBtn');
const auditPaidReconciliationResetBtn = document.getElementById('auditPaidReconciliationResetBtn');
const auditPaidReconciliationFileMeta = document.getElementById('auditPaidReconciliationFileMeta');
const auditPaidReconciliationSummary = document.getElementById('auditPaidReconciliationSummary');
const auditPaidReconciliationMatchedWrap = document.getElementById('auditPaidReconciliationMatchedWrap');
const auditPaidReconciliationMatchedBody = document.getElementById('auditPaidReconciliationMatchedBody');
const auditPaidReconciliationNotSelectedWrap = document.getElementById('auditPaidReconciliationNotSelectedWrap');
const auditPaidReconciliationNotSelectedBody = document.getElementById('auditPaidReconciliationNotSelectedBody');
const auditPaidReconciliationUnmatchedWrap = document.getElementById('auditPaidReconciliationUnmatchedWrap');
const auditPaidReconciliationUnmatchedBody = document.getElementById('auditPaidReconciliationUnmatchedBody');
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
    notification.className = `notification ${type} show`;
    notification.style.display = 'flex';
    clearTimeout(showNotification._timer);
    clearTimeout(showNotification._hideTimer);
    showNotification._timer = setTimeout(() => {
        notification.classList.remove('show');
        showNotification._hideTimer = setTimeout(() => {
            notification.style.display = 'none';
        }, 360);
    }, 3000);
}

function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
}

function normalizeAuditPaidScope(value) {
    return value === 'all' || value === 'others' ? value : 'mine';
}

function getAuditPaidScopeLabel(value) {
    const scope = normalizeAuditPaidScope(value);
    if (scope === 'all') return 'Approved by Me + Others';
    if (scope === 'others') return 'Approved by Others';
    return 'Approved by Me';
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

function normalizeAccountNumber(value) {
    const digits = String(value || '').replace(/\D/g, '');
    return digits.length > 0 && digits.length < 10 ? digits.padStart(10, '0') : digits;
}

function normalizePenNumber(value) {
    return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function normalizeRsaAmount(value) {
    const amount = parseMoney(value);
    return Number.isFinite(amount) ? Math.round(amount * 100) / 100 : 0;
}

function hasRsaAmount(value) {
    return normalizeRsaAmount(value) > 0;
}

function rsaAmountsMatch(left, right) {
    const a = normalizeRsaAmount(left);
    const b = normalizeRsaAmount(right);
    return a > 0 && b > 0 && Math.abs(a - b) < 1;
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
    const uploadedAtMs = getTimestampMillis(getSubmissionOriginalUploadAt(sub));
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

function getCustomerPenNumber(sub = {}) {
    return String(
        sub?.customerDetails?.penNo ||
        sub?.customerDetails?.penNumber ||
        sub?.penNo ||
        sub?.penNumber ||
        ''
    ).trim();
}

function getCustomerPhoneNumber(sub = {}) {
    return String(
        sub?.customerDetails?.phone ||
        sub?.customerDetails?.phoneNumber ||
        sub?.customerDetails?.whatsapp ||
        sub?.customerPhone ||
        sub?.phone ||
        ''
    ).replace(/\D/g, '');
}

function getCustomerNin(sub = {}) {
    return String(
        sub?.customerDetails?.nin ||
        sub?.customerDetails?.nationalId ||
        sub?.customerDetails?.nationalIdentificationNumber ||
        sub?.customerDetails?.customerNIN ||
        sub?.customerNIN ||
        sub?.nin ||
        ''
    ).replace(/\D/g, '');
}

function getSelectedAuditReconciliationFile() {
    return auditReconciliationFileInput?.files?.[0] || null;
}

function getSelectedAuditPaidReconciliationFile() {
    return auditPaidReconciliationFileInput?.files?.[0] || null;
}

function resolveImportColumn(headerMap, acceptedKeys = []) {
    for (const key of acceptedKeys) {
        for (const [colNumber, headerKey] of headerMap.entries()) {
            if (headerKey === key) return colNumber;
        }
    }
    return 0;
}

function getAuditReconciliationColumns(headerMap) {
    return {
        nameColumn: resolveImportColumn(headerMap, ['customername', 'customer', 'name', 'fullname', 'customerfullname', 'accountname', 'membername', 'clientname']),
        accountColumn: resolveImportColumn(headerMap, ['accountnumber', 'accountno', 'acctnumber', 'acctno', 'bankaccountnumber', 'customeraccountnumber', 'customeraccountno']),
        rsaColumn: resolveImportColumn(headerMap, ['25rsabalance', '25percentrsabalance', '25rsabalance', 'twentyfiversabalance', 'rsabalance', 'rsabal', 'retirementsavingsaccountbalance', 'balance', 'amount', 'rsavalue', 'pensionbalance']),
        commissionColumn: resolveImportColumn(headerMap, ['commission', 'commissionamount', 'totalcommissionpayable']),
        rateColumn: resolveImportColumn(headerMap, ['rate', 'commissionrate']),
        uploaderColumn: resolveImportColumn(headerMap, ['uploadername', 'uploader', 'submittedby']),
        agentColumn: resolveImportColumn(headerMap, ['agentname', 'agent']),
        agentAccountColumn: resolveImportColumn(headerMap, ['agentaccountnumber', 'agentaccountno', 'agentacctnumber', 'agentacctno']),
        agentBankColumn: resolveImportColumn(headerMap, ['agentbank', 'agentaccountbank', 'bank']),
        statusColumn: resolveImportColumn(headerMap, ['status', 'paymentstatus', 'auditstatus'])
    };
}

function normalizeAuditReconciliationRow(row = {}) {
    const customerName = String(row.customerName || '').trim();
    const accountNumber = String(row.accountNumber || '').trim();
    const rawRsaBalance = String(row.rawRsaBalance || '').trim();
    const sourceStatus = String(row.sourceStatus || '').trim();
    return {
        ...row,
        customerName,
        accountNumber,
        rawRsaBalance,
        sourceStatus,
        normalizedName: normalizeCustomerName(customerName),
        normalizedAccountNumber: normalizeAccountNumber(accountNumber),
        normalizedRsaBalance: normalizeRsaAmount(rawRsaBalance),
        hasRsaBalance: hasRsaAmount(rawRsaBalance),
        normalizedSourceStatus: sourceStatus.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '')
    };
}

function isClearableImportStatus(row = {}) {
    const status = String(row.normalizedSourceStatus || '').trim();
    return !status || status === 'paid' || status === 'cleared';
}

function parseCsvRows(text = '') {
    const rows = [];
    let row = [];
    let cell = '';
    let inQuotes = false;
    const source = String(text || '');
    for (let i = 0; i < source.length; i += 1) {
        const char = source[i];
        const next = source[i + 1];
        if (char === '"' && inQuotes && next === '"') {
            cell += '"';
            i += 1;
            continue;
        }
        if (char === '"') {
            inQuotes = !inQuotes;
            continue;
        }
        if (char === ',' && !inQuotes) {
            row.push(cell.trim());
            cell = '';
            continue;
        }
        if ((char === '\n' || char === '\r') && !inQuotes) {
            if (char === '\r' && next === '\n') i += 1;
            row.push(cell.trim());
            if (row.some((value) => String(value || '').trim())) rows.push(row);
            row = [];
            cell = '';
            continue;
        }
        cell += char;
    }
    row.push(cell.trim());
    if (row.some((value) => String(value || '').trim())) rows.push(row);
    return rows;
}

function parseAuditReconciliationCsv(text = '') {
    const rows = parseCsvRows(text);
    if (!rows.length) throw new Error('CSV file is empty.');
    const headerMap = new Map();
    rows[0].forEach((header, index) => {
        const normalized = normalizeImportHeader(header);
        if (normalized) headerMap.set(index + 1, normalized);
    });
    const {
        nameColumn,
        accountColumn,
        rsaColumn,
        commissionColumn,
        rateColumn,
        uploaderColumn,
        agentColumn,
        agentAccountColumn,
        agentBankColumn,
        statusColumn
    } = getAuditReconciliationColumns(headerMap);
    if (!nameColumn && !accountColumn) throw new Error('File must contain a name or account number column.');

    const parsedRows = rows.slice(1).map((row, index) => normalizeAuditReconciliationRow({
        rowNumber: index + 2,
        customerName: nameColumn ? getCellText(row[nameColumn - 1]) : '',
        accountNumber: accountColumn ? getCellText(row[accountColumn - 1]) : '',
        rawRsaBalance: rsaColumn ? getCellText(row[rsaColumn - 1]) : '',
        sourceCommission: commissionColumn ? getCellText(row[commissionColumn - 1]) : '',
        sourceRate: rateColumn ? getCellText(row[rateColumn - 1]) : '',
        sourceUploaderName: uploaderColumn ? getCellText(row[uploaderColumn - 1]) : '',
        sourceAgentName: agentColumn ? getCellText(row[agentColumn - 1]) : '',
        sourceAgentAccountNumber: agentAccountColumn ? getCellText(row[agentAccountColumn - 1]) : '',
        sourceAgentBank: agentBankColumn ? getCellText(row[agentBankColumn - 1]) : '',
        sourceStatus: statusColumn ? getCellText(row[statusColumn - 1]) : ''
    })).filter((row) => row.customerName || row.accountNumber || row.rawRsaBalance);

    if (!parsedRows.length) throw new Error('The CSV file does not contain any usable rows.');
    return parsedRows;
}

async function parseAuditReconciliationExcel(file) {
    if (!window.ExcelJS) throw new Error('Excel library is not available right now.');
    const workbook = new window.ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error('No worksheet found in this Excel file.');

    const headerMap = new Map();
    sheet.getRow(1).eachCell((cell, colNumber) => {
        const normalized = normalizeImportHeader(getCellText(cell.value));
        if (normalized) headerMap.set(colNumber, normalized);
    });
    const {
        nameColumn,
        accountColumn,
        rsaColumn,
        commissionColumn,
        rateColumn,
        uploaderColumn,
        agentColumn,
        agentAccountColumn,
        agentBankColumn,
        statusColumn
    } = getAuditReconciliationColumns(headerMap);
    if (!nameColumn && !accountColumn) throw new Error('File must contain a name or account number column.');

    const rows = [];
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        const parsed = normalizeAuditReconciliationRow({
            rowNumber,
            customerName: nameColumn ? getCellText(row.getCell(nameColumn).value) : '',
            accountNumber: accountColumn ? getCellText(row.getCell(accountColumn).value) : '',
            rawRsaBalance: rsaColumn ? getCellText(row.getCell(rsaColumn).value) : '',
            sourceCommission: commissionColumn ? getCellText(row.getCell(commissionColumn).value) : '',
            sourceRate: rateColumn ? getCellText(row.getCell(rateColumn).value) : '',
            sourceUploaderName: uploaderColumn ? getCellText(row.getCell(uploaderColumn).value) : '',
            sourceAgentName: agentColumn ? getCellText(row.getCell(agentColumn).value) : '',
            sourceAgentAccountNumber: agentAccountColumn ? getCellText(row.getCell(agentAccountColumn).value) : '',
            sourceAgentBank: agentBankColumn ? getCellText(row.getCell(agentBankColumn).value) : '',
            sourceStatus: statusColumn ? getCellText(row.getCell(statusColumn).value) : ''
        });
        if (parsed.customerName || parsed.accountNumber || parsed.rawRsaBalance) rows.push(parsed);
    }

    if (!rows.length) throw new Error('The Excel file does not contain any usable rows.');
    return rows;
}

async function parseAuditReconciliationFile(file) {
    if (!file) throw new Error('Select a file first.');
    const name = String(file.name || '').toLowerCase();
    if (name.endsWith('.csv')) return parseAuditReconciliationCsv(await file.text());
    if (!name.endsWith('.xlsx')) throw new Error('Upload an .xlsx or .csv file.');
    return parseAuditReconciliationExcel(file);
}

function mapSubmissionsToAuditReconciliationCandidates(submissions = [], options = {}) {
    const balanceField = options.balanceField === 'twentyFive' ? 'twentyFive' : 'rsaBalance';
    return submissions.map((sub) => {
        const financials = getSubmissionFinancials(sub);
        const matchBalance = balanceField === 'twentyFive' ? financials.twentyFive : financials.rsaBalance;
        const accountNumber = getCustomerAccountNumber(sub);
        const customerName = String(sub.customerName || sub?.customerDetails?.name || '').trim();
        return {
            submission: sub,
            customerName,
            normalizedName: normalizeCustomerName(customerName),
            accountNumber,
            normalizedAccountNumber: normalizeAccountNumber(accountNumber),
            rsaBalance: matchBalance,
            fullRsaBalance: financials.rsaBalance,
            twentyFiveRsaBalance: financials.twentyFive,
            normalizedRsaBalance: normalizeRsaAmount(matchBalance)
        };
    });
}

function getAuditReconciliationCandidates() {
    return mapSubmissionsToAuditReconciliationCandidates(allSubmissions);
}

function getAuditPaidReconciliationCandidates(scope = currentAuditPaidScope) {
    const resolvedScope = normalizeAuditPaidScope(scope);
    const currentEmail = normalizeEmail(currentUser?.email);
    const rows = allSubmissions
        .filter((sub) => ['paid', 'cleared'].includes(String(sub.status || '').toLowerCase()))
        .filter((sub) => {
            if (resolvedScope === 'all') return true;
            const approvedBy = getAuditApprovalEmail(sub);
            if (!currentEmail) return resolvedScope !== 'mine';
            return resolvedScope === 'mine' ? approvedBy === currentEmail : approvedBy !== currentEmail;
        });
    return mapSubmissionsToAuditReconciliationCandidates(rows, { balanceField: 'twentyFive' });
}

function isSystemPaidForClearing(row = {}) {
    return String(row?.submission?.status || '').toLowerCase() === 'paid';
}

function isAuditPaidRowClearable(row = {}) {
    return isSystemPaidForClearing(row) && isClearableImportStatus(row);
}

function getLatestSubmissionMillis(sub = {}) {
    return getTimestampMillis(sub.updatedAt || sub.paidAt || sub.clearedAt || sub.uploadedAt || getSubmissionCurrentStageEntryAt(sub));
}

function chooseAuditReconciliationMatch(row, candidates = []) {
    const accountMatches = row.normalizedAccountNumber
        ? candidates.filter((candidate) => candidate.normalizedAccountNumber === row.normalizedAccountNumber)
        : [];
    const nameMatches = row.normalizedName
        ? candidates.filter((candidate) => candidate.normalizedName === row.normalizedName)
        : [];
    const pickLatest = (items = []) => items.slice().sort((a, b) => getLatestSubmissionMillis(b.submission) - getLatestSubmissionMillis(a.submission))[0] || null;

    if (row.hasRsaBalance) {
        const accountBalanceMatch = pickLatest(accountMatches.filter((candidate) => rsaAmountsMatch(row.normalizedRsaBalance, candidate.normalizedRsaBalance)));
        if (accountBalanceMatch) return { candidate: accountBalanceMatch, method: 'Account + RSA balance', balanceMatches: true };
    }

    const accountMatch = pickLatest(accountMatches);
    if (accountMatch) return { candidate: accountMatch, method: 'Account number', balanceMatches: row.hasRsaBalance ? rsaAmountsMatch(row.normalizedRsaBalance, accountMatch.normalizedRsaBalance) : true };

    if (row.hasRsaBalance) {
        const nameBalanceMatch = pickLatest(nameMatches.filter((candidate) => rsaAmountsMatch(row.normalizedRsaBalance, candidate.normalizedRsaBalance)));
        if (nameBalanceMatch) return { candidate: nameBalanceMatch, method: 'Name + RSA balance', balanceMatches: true };
    }

    const nameMatch = pickLatest(nameMatches);
    if (nameMatch) return { candidate: nameMatch, method: 'Name', balanceMatches: row.hasRsaBalance ? rsaAmountsMatch(row.normalizedRsaBalance, nameMatch.normalizedRsaBalance) : true };

    return null;
}

function computeAuditReconciliationResultFromCandidates(rows = [], candidates = []) {
    const matchedRows = [];
    const unmatchedRows = [];

    rows.forEach((row) => {
        const match = chooseAuditReconciliationMatch(row, candidates);
        if (!match) {
            unmatchedRows.push({
                ...row,
                reason: 'Name/account number was not found in system records.'
            });
            return;
        }

        const { candidate, method, balanceMatches } = match;
        matchedRows.push({
            ...row,
            matchMethod: method,
            balanceMatches,
            submission: candidate.submission,
            submissionId: candidate.submission.id,
            systemName: candidate.customerName || candidate.submission.customerName || '-',
            uploaderName: getUserDisplayName(candidate.submission.uploadedBy || candidate.submission.auditCommissionSubmittedBy || ''),
            systemAccountNumber: candidate.accountNumber,
            systemRsaBalance: candidate.rsaBalance,
            systemStatus: statusLabel(candidate.submission.status || '-')
        });
    });

    return {
        totalRows: rows.length,
        matchedRows,
        unmatchedRows,
        exactCount: matchedRows.filter((row) => row.balanceMatches).length,
        balanceDifferenceCount: matchedRows.filter((row) => !row.balanceMatches).length,
        systemCount: candidates.length
    };
}

function computeAuditReconciliationResult(rows = []) {
    return computeAuditReconciliationResultFromCandidates(rows, getAuditReconciliationCandidates());
}

function getAuditableDuplicateSubmissions() {
    const ignoredStatuses = new Set(['draft', 'rejected', 'rejected_by_reviewer', 'rejected_by_rsa', 'deleted']);
    return allSubmissions.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        if (sub.auditDuplicateIgnored === true) return false;
        if (sub.auditDuplicateRejected === true || sub.auditDuplicateDeleted === true || sub.auditDuplicateCorrectionStatus === 'corrected') return false;
        return !ignoredStatuses.has(status);
    });
}

function isAuditDuplicateHandled(sub = {}) {
    return sub.auditDuplicateIgnored === true ||
        sub.auditDuplicateRejected === true ||
        sub.auditDuplicateDeleted === true ||
        String(sub.auditDuplicateCorrectionStatus || '').trim().toLowerCase() === 'corrected';
}

function isAuditDuplicateCorrectionPending(sub = {}) {
    return String(sub.status || '').trim().toLowerCase() === 'audit_pending'
        || String(sub.auditDuplicateCorrectionStatus || '').trim().toLowerCase() === 'pending';
}

function getAuditDuplicatePreviousStatus(sub = {}) {
    return String(sub.auditDuplicatePreviousStatus || '').trim().toLowerCase();
}

function inferAuditDuplicateRestoreStatus(sub = {}) {
    const status = String(sub.status || '').trim().toLowerCase();
    if (status === 'cleared' || sub.clearedAt || sub.auditClearedAt || sub.clearedBy) return 'cleared';
    if (status === 'paid' || sub.paidAt || sub.paidBy || sub.auditCommissionAcceptedAt || sub.auditCommissionAcceptedBy) return 'paid';
    if (
        status === 'sent_to_pfa' ||
        status === 'rsa_submitted' ||
        sub.finalSubmitted === true ||
        sub.rsaSubmitted === true ||
        sub.finalSubmittedAt ||
        sub.rsaSubmittedAt ||
        sub.paymentAssignedAt ||
        sub.assignedToPayment
    ) {
        return 'sent_to_pfa';
    }
    if (status === 'processing_to_pfa' || status === 'approved' || sub.rsaReady === true || sub.assignedToRSA) {
        return 'processing_to_pfa';
    }
    return '';
}

function shouldRestoreAuditDuplicateCorrection(sub = {}) {
    const restoredStatus = inferAuditDuplicateRestoreStatus(sub) || getAuditDuplicatePreviousStatus(sub);
    if (!restoredStatus) return false;
    if (sub.auditDuplicateRejected !== true && !sub.auditDuplicateRejectionReason) return false;
    const currentStatus = String(sub.status || '').trim().toLowerCase();
    const hasCorrectionSignal = isAuditDuplicateCorrectionPending(sub) || Boolean(sub.auditDuplicateResubmittedAt);
    if (!hasCorrectionSignal) return false;
    if (isAuditDuplicateCorrectionPending(sub)) return true;
    return currentStatus !== restoredStatus && currentStatus === 'pending';
}

function restoreAuditDuplicateCorrections(rows = []) {
    rows.filter(shouldRestoreAuditDuplicateCorrection).forEach((sub) => {
        const submissionId = String(sub.id || '').trim();
        const restoredStatus = inferAuditDuplicateRestoreStatus(sub) || getAuditDuplicatePreviousStatus(sub);
        if (!submissionId || !restoredStatus || auditDuplicateRestoreInFlight.has(submissionId)) return;
        auditDuplicateRestoreInFlight.add(submissionId);
        updateDoc(doc(db, 'submissions', submissionId), {
            status: restoredStatus,
            auditDuplicateCorrectionStatus: 'corrected',
            auditDuplicateRestoredStatus: restoredStatus,
            auditDuplicateAutoRestoredAt: serverTimestamp(),
            latestRejectedStage: '',
            updatedAt: serverTimestamp()
        }).catch(() => {}).finally(() => {
            auditDuplicateRestoreInFlight.delete(submissionId);
        });
    });
}

function isAuditDuplicateIgnoredCorrectionCopy(sub = {}) {
    const status = String(sub.status || '').trim().toLowerCase();
    return ['rejected', 'rejected_by_reviewer', 'rejected_by_rsa'].includes(status);
}

function getDuplicateLineageKey(sub = {}) {
    const explicit = String(
        sub.rootSubmissionId ||
        sub.originalSubmissionId ||
        sub.parentSubmissionId ||
        sub.sourceSubmissionId ||
        sub.previousSubmissionId ||
        sub.linkedSubmissionId ||
        sub.duplicateOf ||
        ''
    ).trim();
    return explicit || String(sub.id || '').trim();
}

function getDuplicateRepresentative(rows = []) {
    return rows
        .slice()
        .sort((a, b) => {
            const aRejected = isAuditDuplicateIgnoredCorrectionCopy(a) ? 1 : 0;
            const bRejected = isAuditDuplicateIgnoredCorrectionCopy(b) ? 1 : 0;
            if (aRejected !== bRejected) return aRejected - bRejected;
            return getTimestampMillis(getSubmissionOriginalUploadAt(b)) - getTimestampMillis(getSubmissionOriginalUploadAt(a));
        })[0] || rows[0] || null;
}

function collapseDuplicateRowsByLineage(rows = []) {
    const byLineage = new Map();
    rows.forEach((sub) => {
        const lineageKey = getDuplicateLineageKey(sub);
        if (!lineageKey) return;
        const next = byLineage.get(lineageKey) || [];
        next.push(sub);
        byLineage.set(lineageKey, next);
    });
    return Array.from(byLineage.entries())
        .map(([lineageKey, lineageRows]) => ({
            lineageKey,
            rowCount: lineageRows.length,
            representative: getDuplicateRepresentative(lineageRows)
        }))
        .filter((item) => item.representative);
}

function buildAuditDuplicateScanResult() {
    const buckets = new Map();
    const addBucket = (key, signal, strength, sub) => {
        if (!key) return;
        const bucketKey = `${signal}:${key}`;
        const existing = buckets.get(bucketKey) || { key, signal, strength, rows: [] };
        if (!existing.rows.some((row) => row.id === sub.id)) existing.rows.push(sub);
        buckets.set(bucketKey, existing);
    };

    const ignoredCorrectionCount = allSubmissions.filter(isAuditDuplicateIgnoredCorrectionCopy).length;
    const handledDuplicateCount = allSubmissions.filter(isAuditDuplicateHandled).length;
    const auditableRows = getAuditableDuplicateSubmissions();
    auditableRows.forEach((sub) => {
        const accountKey = normalizeAccountNumber(getCustomerAccountNumber(sub));
        const penKey = normalizePenNumber(getCustomerPenNumber(sub));
        const nameKey = normalizeCustomerName(sub.customerName || sub?.customerDetails?.name || '');
        const phoneKey = getCustomerPhoneNumber(sub);
        const ninKey = getCustomerNin(sub);
        const pfaKey = normalizeCustomerName(getSubmissionPfaName(sub));
        if (accountKey) addBucket(accountKey, 'Account Number', 'strong', sub);
        if (penKey) addBucket(penKey, 'PEN', 'strong', sub);
        if (phoneKey && phoneKey.length >= 10) addBucket(phoneKey, 'Phone Number', 'strong', sub);
        if (ninKey && ninKey.length >= 10) addBucket(ninKey, 'NIN', 'strong', sub);
        if (nameKey.length >= 8 && nameKey.split(' ').length >= 2 && pfaKey) {
            addBucket(`${nameKey}|${pfaKey}`, 'Customer Name + PFA', 'possible', sub);
        }
    });

    const groupMap = new Map();
    Array.from(buckets.values()).forEach((bucket) => {
        const independentRows = collapseDuplicateRowsByLineage(bucket.rows);
        if (independentRows.length <= 1) return;
        const rows = independentRows
            .map((item) => item.representative)
            .sort((a, b) => getTimestampMillis(getSubmissionOriginalUploadAt(a)) - getTimestampMillis(getSubmissionOriginalUploadAt(b)));
        const clusterKey = independentRows.map((item) => item.lineageKey).sort().join('|');
        const existing = groupMap.get(clusterKey) || {
            key: '',
            signals: [],
            strength: 'possible',
            rows,
            hiddenCorrectionRows: 0
        };
        existing.signals.push({ signal: bucket.signal, key: bucket.key, strength: bucket.strength });
        existing.strength = existing.strength === 'strong' || bucket.strength === 'strong' ? 'strong' : 'possible';
        existing.rows = rows;
        existing.hiddenCorrectionRows += independentRows.reduce((sum, item) => sum + Math.max(0, Number(item.rowCount || 0) - 1), 0);
        groupMap.set(clusterKey, existing);
    });

    const groups = Array.from(groupMap.values())
        .map((group) => ({
            ...group,
            signal: group.signals.map((item) => item.signal).join(' + '),
            key: group.signals.map((item) => `${item.signal}: ${item.key}`).join('; ')
        }))
        .sort((a, b) => {
            if (a.strength !== b.strength) return a.strength === 'strong' ? -1 : 1;
            return b.rows.length - a.rows.length;
        });

    const duplicateApplicationIds = new Set();
    groups.forEach((group) => group.rows.forEach((sub) => duplicateApplicationIds.add(sub.id)));
    return {
        scannedCount: auditableRows.length,
        ignoredCorrectionCount,
        handledDuplicateCount,
        groups,
        duplicateCount: duplicateApplicationIds.size,
        strongGroupCount: groups.filter((group) => group.strength === 'strong').length,
        possibleGroupCount: groups.filter((group) => group.strength !== 'strong').length
    };
}

function recomputeAuditPaidReconciliationResult() {
    if (!auditPaidReconciliationSourceRows.length) {
        auditPaidReconciliationResult = null;
        auditPaidReconciliationSelectedSubmissionIds = [];
        return;
    }
    auditPaidReconciliationResult = computeAuditReconciliationResultFromCandidates(
        auditPaidReconciliationSourceRows,
        getAuditPaidReconciliationCandidates(currentAuditPaidScope)
    );
    const validSubmissionIds = new Set((auditPaidReconciliationResult?.matchedRows || []).filter((row) => isAuditPaidRowClearable(row)).map((row) => row.submissionId).filter(Boolean));
    auditPaidReconciliationSelectedSubmissionIds = auditPaidReconciliationSelectedSubmissionIds.filter((submissionId) => validSubmissionIds.has(submissionId));
}

function resetAuditPaidReconciliationState() {
    auditPaidReconciliationSourceRows = [];
    auditPaidReconciliationResult = null;
    auditPaidReconciliationFileName = '';
    auditPaidReconciliationSelectedSubmissionIds = [];
    auditPaidReconciliationActiveResultsTab = 'matched';
    if (auditPaidReconciliationFileInput) auditPaidReconciliationFileInput.value = '';
}

function openAuditPaidReconciliationUploadModal() {
    renderAuditPaidReconciliation();
    auditPaidReconciliationUploadModal?.classList.add('active');
}

function closeAuditPaidReconciliationUploadModal() {
    auditPaidReconciliationUploadModal?.classList.remove('active');
}

function openAuditPaidReconciliationResultsModal() {
    if (!auditPaidReconciliationResult) return;
    renderAuditPaidReconciliation();
    auditPaidReconciliationResultsModal?.classList.add('active');
}

function closeAuditPaidReconciliationResultsModal() {
    auditPaidReconciliationResultsModal?.classList.remove('active');
}

function showAuditPaidReconciliationResultsTab(tabName = 'matched') {
    auditPaidReconciliationActiveResultsTab = ['matched', 'not-selected', 'unmatched'].includes(tabName) ? tabName : 'matched';
    document.querySelectorAll('[data-paid-reconciliation-view]').forEach((button) => {
        const isActive = button.dataset.paidReconciliationView === auditPaidReconciliationActiveResultsTab;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    });
    renderAuditPaidReconciliation();
}

function resetAuditReconciliationState() {
    auditReconciliationSourceRows = [];
    auditReconciliationResult = null;
    auditReconciliationFileName = '';
    auditDuplicateScanResult = null;
    auditReconciliationActiveView = 'excel';
    if (auditReconciliationFileInput) auditReconciliationFileInput.value = '';
    renderAuditReconciliation();
    renderAuditDuplicateScan();
}

function syncAuditReconciliationViewVisibility() {
    const activeView = AUDIT_RECONCILIATION_VIEWS.includes(auditReconciliationActiveView) ? auditReconciliationActiveView : 'excel';
    document.querySelectorAll('[data-audit-reconciliation-view]').forEach((button) => {
        const isActive = button.dataset.auditReconciliationView === activeView;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    });
    document.querySelectorAll('[data-audit-reconciliation-panel]').forEach((panel) => {
        if (panel.dataset.auditReconciliationPanel !== activeView) {
            panel.style.display = 'none';
        }
    });
}

function showAuditReconciliationView(viewName = 'excel') {
    auditReconciliationActiveView = AUDIT_RECONCILIATION_VIEWS.includes(viewName) ? viewName : 'excel';
    renderAuditReconciliation();
    renderAuditDuplicateScan();
    renderAuditDuplicateIgnoredTable();
    renderAuditDuplicateRejectedTable();
    syncAuditReconciliationViewVisibility();
}

function renderAuditPaidReconciliation() {
    const selectedFile = getSelectedAuditPaidReconciliationFile();
    const scopeLabel = getAuditPaidScopeLabel(currentAuditPaidScope);
    const matchedRows = auditPaidReconciliationResult?.matchedRows || [];
    const unmatchedRows = auditPaidReconciliationResult?.unmatchedRows || [];
    const selectedSubmissionIds = Array.from(new Set(auditPaidReconciliationSelectedSubmissionIds.filter(Boolean)));
    const selectedSet = new Set(selectedSubmissionIds);
    const countedSelectedIds = new Set();
    const notSelectedRows = matchedRows.reduce((rows, row) => {
        const submissionId = String(row.submissionId || '').trim();
        const isSelected = submissionId && selectedSet.has(submissionId);
        const isClearable = isAuditPaidRowClearable(row);
        if (!isSelected || !isClearable) {
            rows.push({
                ...row,
                notSelectedReason: !isClearable
                    ? (!isSystemPaidForClearing(row)
                        ? `System status is ${row.systemStatus || 'not paid'}`
                        : `Uploaded status is ${row.sourceStatus || 'not clearable'}`)
                    : 'Manually not selected'
            });
            return rows;
        }
        if (countedSelectedIds.has(submissionId)) {
            rows.push({
                ...row,
                notSelectedReason: 'Duplicate upload row for the same application'
            });
            return rows;
        }
        countedSelectedIds.add(submissionId);
        return rows;
    }, []);
    const selectedCount = selectedSubmissionIds.length;

    if (auditPaidReconciliationFileMeta) {
        auditPaidReconciliationFileMeta.textContent = auditPaidReconciliationFileName
            ? `Selected file: ${auditPaidReconciliationFileName}. Matching against ${scopeLabel}.`
            : `No file selected. Matching runs against ${scopeLabel}.`;
    }
    if (auditPaidReconciliationRunBtn) auditPaidReconciliationRunBtn.disabled = !selectedFile;
    if (auditPaidReconciliationResetBtn) auditPaidReconciliationResetBtn.disabled = !selectedFile && !auditPaidReconciliationResult && !auditPaidReconciliationFileName;
    if (auditPaidReconciliationSelectAllBtn) auditPaidReconciliationSelectAllBtn.disabled = !matchedRows.some((row) => isAuditPaidRowClearable(row));
    if (auditPaidReconciliationClearSelectedBtn) auditPaidReconciliationClearSelectedBtn.disabled = !selectedCount;

    if (!auditPaidReconciliationResult) {
        if (auditPaidReconciliationSummary) {
            auditPaidReconciliationSummary.style.display = 'none';
            auditPaidReconciliationSummary.innerHTML = '';
        }
        if (auditPaidReconciliationMatchedWrap) auditPaidReconciliationMatchedWrap.style.display = 'none';
        if (auditPaidReconciliationMatchedBody) auditPaidReconciliationMatchedBody.innerHTML = '';
        if (auditPaidReconciliationNotSelectedWrap) auditPaidReconciliationNotSelectedWrap.style.display = 'none';
        if (auditPaidReconciliationNotSelectedBody) auditPaidReconciliationNotSelectedBody.innerHTML = '';
        if (auditPaidReconciliationUnmatchedWrap) auditPaidReconciliationUnmatchedWrap.style.display = 'none';
        if (auditPaidReconciliationUnmatchedBody) auditPaidReconciliationUnmatchedBody.innerHTML = '';
        return;
    }

    if (auditPaidReconciliationSummary) {
        auditPaidReconciliationSummary.style.display = 'grid';
        auditPaidReconciliationSummary.innerHTML = [
            { label: 'Uploaded Rows', value: auditPaidReconciliationResult.totalRows },
            { label: 'Found in System', value: matchedRows.length },
            { label: 'Exact Balance', value: auditPaidReconciliationResult.exactCount },
            { label: 'Balance Differs', value: auditPaidReconciliationResult.balanceDifferenceCount },
            { label: 'Not Found', value: unmatchedRows.length },
            { label: 'Not Selected', value: notSelectedRows.length },
            { label: 'Selected for Clearing', value: selectedCount }
        ].map((chip) => `
            <div class="audit-reconciliation-chip">
                <span>${escapeHtml(chip.label)}</span>
                <strong>${String(chip.value)}</strong>
            </div>
        `).join('');
    }

    document.querySelectorAll('[data-paid-reconciliation-view]').forEach((button) => {
        const isActive = button.dataset.paidReconciliationView === auditPaidReconciliationActiveResultsTab;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    });

    if (auditPaidReconciliationMatchedWrap) auditPaidReconciliationMatchedWrap.style.display = auditPaidReconciliationActiveResultsTab === 'matched' ? 'block' : 'none';
    if (auditPaidReconciliationMatchedBody) {
        auditPaidReconciliationMatchedBody.innerHTML = matchedRows.length
            ? matchedRows.map((row) => {
                const isPaid = isSystemPaidForClearing(row);
                const isClearable = isAuditPaidRowClearable(row);
                const statusClass = isPaid && row.balanceMatches ? 'audit-recon-status exact' : 'audit-recon-status partial';
                const matchText = !isPaid
                    ? `Found but ${row.systemStatus || 'not paid'}`
                    : (!isClearableImportStatus(row)
                        ? `Found but uploaded status is ${row.sourceStatus || 'not clearable'}`
                        : (row.balanceMatches ? `Exact match (${row.matchMethod})` : `Balance differs (${row.matchMethod})`));
                const isChecked = selectedSet.has(row.submissionId) ? 'checked' : '';
                return `
                    <tr>
                        <td><input type="checkbox" data-paid-recon-select="${escapeHtml(row.submissionId)}" ${isChecked} ${isClearable ? '' : 'disabled'}></td>
                        <td>${escapeHtml(row.rowNumber)}</td>
                        <td>${escapeHtml(row.customerName || '-')}</td>
                        <td>${escapeHtml(row.accountNumber || '-')}</td>
                        <td>${row.hasRsaBalance ? formatCurrency(row.normalizedRsaBalance) : '-'}</td>
                        <td>${escapeHtml(row.sourceStatus || '-')}</td>
                        <td><strong>${escapeHtml(row.systemName || '-')}</strong></td>
                        <td>${escapeHtml(row.uploaderName || '-')}</td>
                        <td>${escapeHtml(row.systemAccountNumber || '-')}</td>
                        <td>${escapeHtml(formatDate(row.submission?.paidAt || row.submission?.updatedAt || ''))}</td>
                        <td><span class="${isClearable ? statusClass : 'audit-recon-status partial'}">${escapeHtml(matchText)}</span></td>
                        <td><button type="button" class="action-btn" onclick="window.openMonitoringApplicationDetails('${row.submissionId}')"><i class="fas fa-eye"></i> View</button></td>
                    </tr>
                `;
            }).join('')
            : '<tr><td colspan="12" class="no-data">No paid applications matched this file</td></tr>';
    }

    if (auditPaidReconciliationNotSelectedWrap) auditPaidReconciliationNotSelectedWrap.style.display = auditPaidReconciliationActiveResultsTab === 'not-selected' ? 'block' : 'none';
    if (auditPaidReconciliationNotSelectedBody) {
        auditPaidReconciliationNotSelectedBody.innerHTML = notSelectedRows.length
            ? notSelectedRows.map((row) => {
                const reason = row.notSelectedReason || 'Not selected for clearing';
                return `
                    <tr>
                        <td>${escapeHtml(row.rowNumber)}</td>
                        <td>${escapeHtml(row.customerName || '-')}</td>
                        <td>${escapeHtml(row.accountNumber || '-')}</td>
                        <td>${row.hasRsaBalance ? formatCurrency(row.normalizedRsaBalance) : '-'}</td>
                        <td>${escapeHtml(row.sourceStatus || '-')}</td>
                        <td><strong>${escapeHtml(row.systemName || '-')}</strong></td>
                        <td>${escapeHtml(row.systemAccountNumber || '-')}</td>
                        <td>${escapeHtml(row.systemStatus || '-')}</td>
                        <td><span class="audit-recon-status partial">${escapeHtml(reason)}</span></td>
                        <td><button type="button" class="action-btn" onclick="window.openMonitoringApplicationDetails('${row.submissionId}')"><i class="fas fa-eye"></i> View</button></td>
                    </tr>
                `;
            }).join('')
            : '<tr><td colspan="10" class="no-data">All matched clearable records are selected</td></tr>';
    }

    if (auditPaidReconciliationUnmatchedWrap) auditPaidReconciliationUnmatchedWrap.style.display = auditPaidReconciliationActiveResultsTab === 'unmatched' ? 'block' : 'none';
    if (auditPaidReconciliationUnmatchedBody) {
        auditPaidReconciliationUnmatchedBody.innerHTML = unmatchedRows.length
            ? unmatchedRows.map((row) => `
                <tr>
                    <td>${escapeHtml(row.rowNumber)}</td>
                    <td>${escapeHtml(row.customerName || '-')}</td>
                    <td>${escapeHtml(row.accountNumber || '-')}</td>
                    <td>${row.hasRsaBalance ? formatCurrency(row.normalizedRsaBalance) : '-'}</td>
                    <td>${escapeHtml(row.reason || 'No paid match found')}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="5" class="no-data">No missing paid records</td></tr>';
    }
}

function renderAuditDuplicateScan() {
    if (!auditDuplicateScanResult) auditDuplicateScanResult = buildAuditDuplicateScanResult();
    const selectedFile = getSelectedAuditReconciliationFile();
    if (auditReconciliationClearBtn) auditReconciliationClearBtn.disabled = !selectedFile && !auditReconciliationResult && !auditDuplicateScanResult;
    if (auditReconciliationExportBtn) auditReconciliationExportBtn.disabled = !auditReconciliationResult && !auditDuplicateScanResult;
    if (!auditDuplicateScanResult) {
        if (auditDuplicateScanSummary) {
            auditDuplicateScanSummary.style.display = 'none';
            auditDuplicateScanSummary.innerHTML = '';
        }
        if (auditDuplicateScanWrap) auditDuplicateScanWrap.style.display = 'none';
        if (auditDuplicateScanBody) auditDuplicateScanBody.innerHTML = '';
        syncAuditReconciliationViewVisibility();
        return;
    }

    const groups = auditDuplicateScanResult.groups || [];
    if (auditDuplicateScanSummary) {
        auditDuplicateScanSummary.style.display = 'grid';
        auditDuplicateScanSummary.innerHTML = [
            { label: 'System Records Scanned', value: auditDuplicateScanResult.scannedCount },
            { label: 'Rejected Corrections Ignored', value: auditDuplicateScanResult.ignoredCorrectionCount || 0 },
            { label: 'Handled History', value: auditDuplicateScanResult.handledDuplicateCount || 0 },
            { label: 'Duplicate Groups', value: groups.length },
            { label: 'Applications Affected', value: auditDuplicateScanResult.duplicateCount },
            { label: 'Strong Signals', value: auditDuplicateScanResult.strongGroupCount },
            { label: 'Possible Matches', value: auditDuplicateScanResult.possibleGroupCount }
        ].map((chip) => `
            <div class="audit-reconciliation-chip">
                <span>${escapeHtml(chip.label)}</span>
                <strong>${String(chip.value)}</strong>
            </div>
        `).join('');
    }

    if (auditDuplicateScanWrap) auditDuplicateScanWrap.style.display = 'block';
    if (!auditDuplicateScanBody) return;
    if (!groups.length) {
        auditDuplicateScanBody.innerHTML = '<tr><td colspan="7" class="no-data">No duplicate applications found</td></tr>';
        syncAuditReconciliationViewVisibility();
        return;
    }

    auditDuplicateScanBody.innerHTML = groups.map((group, groupIndex) => {
        const signalClass = group.strength === 'strong' ? 'audit-recon-status partial' : 'audit-recon-status info';
        const groupRows = group.rows.map((sub) => {
            const accountNumber = getCustomerAccountNumber(sub);
            const penNumber = getCustomerPenNumber(sub);
            return `
                <tr class="audit-duplicate-member-row">
                    <td class="audit-duplicate-customer"><strong>${escapeHtml(sub.customerName || sub?.customerDetails?.name || 'Unknown')}</strong></td>
                    <td class="audit-duplicate-nowrap">${escapeHtml(accountNumber || '-')}</td>
                    <td class="audit-duplicate-nowrap">${escapeHtml(penNumber || '-')}</td>
                    <td><span class="audit-duplicate-status">${escapeHtml(statusLabel(sub.status || '-'))}</span></td>
                    <td>${escapeHtml(getUserDisplayName(sub.uploadedBy || ''))}</td>
                    <td class="audit-duplicate-nowrap">${escapeHtml(formatDate(getSubmissionOriginalUploadAt(sub)))}</td>
                    <td>
                        <div class="audit-duplicate-actions">
                            <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
                            <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationTrack('${sub.id}')"><i class="fas fa-route"></i> Track</button>
                            <button type="button" class="action-btn audit-duplicate-ignore-btn" onclick="window.ignoreAuditDuplicateApplication('${sub.id}')"><i class="fas fa-eye-slash"></i> Ignore</button>
                            <button type="button" class="action-btn audit-duplicate-reject-btn" onclick="window.rejectAuditDuplicateApplication('${sub.id}')"><i class="fas fa-ban"></i> Reject</button>
                            <button type="button" class="action-btn audit-duplicate-delete-btn" onclick="window.deleteAuditDuplicateApplication('${sub.id}')"><i class="fas fa-trash"></i> Delete</button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
        const spacer = groupIndex < groups.length - 1 ? '<tr class="audit-duplicate-group-spacer"><td colspan="7"></td></tr>' : '';
        return `
            <tr class="audit-duplicate-group-row">
                <td colspan="7">
                    <div class="audit-duplicate-group-heading">
                        <span class="${signalClass} audit-duplicate-group-badge">${escapeHtml(group.strength === 'strong' ? 'Strong duplicate signal' : 'Possible duplicate signal')}</span>
                        <span class="audit-duplicate-group-count">${escapeHtml(`${group.rows.length} applications`)}</span>
                        <strong>Group ${groupIndex + 1}</strong>
                        <span class="audit-duplicate-group-signal">${escapeHtml(group.key || group.signal || 'Duplicate match')}</span>
                        <button type="button" class="action-btn audit-duplicate-reject-btn audit-duplicate-group-reject-btn" onclick="window.rejectAuditDuplicateGroup(${groupIndex})"><i class="fas fa-ban"></i> Reject Group</button>
                    </div>
                </td>
            </tr>
            ${groupRows}
            ${spacer}
        `;
    }).join('');
    syncAuditReconciliationViewVisibility();
}

function getAuditDuplicateIgnoredRows() {
    return allSubmissions
        .filter((sub) => sub.auditDuplicateIgnored === true)
        .sort((a, b) => getTimestampMillis(b.auditDuplicateIgnoredAt || b.updatedAt) - getTimestampMillis(a.auditDuplicateIgnoredAt || a.updatedAt));
}

function normalizeAuditDuplicateHistoryFilter(value) {
    const normalized = String(value || '').trim().toLowerCase();
    return ['corrected', 'awaiting'].includes(normalized) ? normalized : 'all';
}

function getAuditDuplicateCorrectionState(sub = {}) {
    const correctionStatus = String(sub.auditDuplicateCorrectionStatus || '').trim().toLowerCase();
    const correctedAt = sub.auditDuplicateResubmittedAt || null;
    return correctedAt || (correctionStatus === 'corrected' && sub.auditDuplicateResubmittedBy) ? 'corrected' : 'awaiting';
}

function getAuditDuplicateRejectedRows() {
    const filter = normalizeAuditDuplicateHistoryFilter(currentAuditDuplicateHistoryFilter);
    return allSubmissions
        .filter((sub) => sub.auditDuplicateRejected === true)
        .filter((sub) => filter === 'all' || getAuditDuplicateCorrectionState(sub) === filter)
        .sort((a, b) => getTimestampMillis(b.auditDuplicateRejectedAt || b.latestRejectedAt || b.updatedAt) - getTimestampMillis(a.auditDuplicateRejectedAt || a.latestRejectedAt || a.updatedAt));
}

function getAuditDuplicateCorrectionLabel(sub = {}) {
    const correctionStatus = String(sub.auditDuplicateCorrectionStatus || '').trim().toLowerCase();
    const correctedAt = sub.auditDuplicateResubmittedAt || null;
    if (correctedAt || (correctionStatus === 'corrected' && sub.auditDuplicateResubmittedBy)) {
        const correctedBy = getUserDisplayName(sub.auditDuplicateResubmittedBy || sub.uploadedBy || '');
        const correctedAtText = correctedAt ? formatDate(correctedAt) : '-';
        return `Corrected by ${correctedBy || '-'} at ${correctedAtText}`;
    }
    return 'Awaiting uploader correction';
}

function getAuditDuplicateReturnedToLabel(sub = {}) {
    const restoredStatus = inferAuditDuplicateRestoreStatus(sub)
        || String(sub.auditDuplicateRestoredStatus || '').trim()
        || String(sub.auditDuplicatePreviousStatus || '').trim()
        || String(sub.status || '').trim();
    return statusLabel(restoredStatus || '-');
}

function renderAuditDuplicateIgnoredTable() {
    if (!auditDuplicateIgnoredWrap || !auditDuplicateIgnoredBody) return;
    const rows = getAuditDuplicateIgnoredRows();
    auditDuplicateIgnoredWrap.style.display = 'block';
    auditDuplicateIgnoredBody.innerHTML = rows.length
        ? rows.map((sub) => `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || sub?.customerDetails?.name || 'Unknown')}</strong></td>
                <td>${escapeHtml(getCustomerAccountNumber(sub) || '-')}</td>
                <td>${escapeHtml(getCustomerPenNumber(sub) || '-')}</td>
                <td>${escapeHtml(getUserDisplayName(sub.uploadedBy || ''))}</td>
                <td>${escapeHtml(getUserDisplayName(sub.auditDuplicateIgnoredBy || ''))}</td>
                <td>${escapeHtml(formatDate(sub.auditDuplicateIgnoredAt || sub.updatedAt))}</td>
                <td>
                    <div class="audit-duplicate-actions">
                        <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
                        <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationTrack('${sub.id}')"><i class="fas fa-route"></i> Track</button>
                    </div>
                </td>
            </tr>
        `).join('')
        : '<tr><td colspan="7" class="no-data">No ignored duplicate applications</td></tr>';
    syncAuditReconciliationViewVisibility();
}

function renderAuditDuplicateRejectedTable() {
    if (!auditDuplicateRejectedWrap || !auditDuplicateRejectedBody) return;
    currentAuditDuplicateHistoryFilter = normalizeAuditDuplicateHistoryFilter(currentAuditDuplicateHistoryFilter);
    if (auditDuplicateHistoryFilter) auditDuplicateHistoryFilter.value = currentAuditDuplicateHistoryFilter;
    const rows = getAuditDuplicateRejectedRows();
    const emptyMessage = currentAuditDuplicateHistoryFilter === 'corrected'
        ? 'No corrected duplicate correction history'
        : currentAuditDuplicateHistoryFilter === 'awaiting'
            ? 'No duplicate applications awaiting correction'
            : 'No duplicate correction history';
    auditDuplicateRejectedWrap.style.display = 'block';
    auditDuplicateRejectedBody.innerHTML = rows.length
        ? rows.map((sub) => `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || sub?.customerDetails?.name || 'Unknown')}</strong></td>
                <td>${escapeHtml(getCustomerAccountNumber(sub) || '-')}</td>
                <td>${escapeHtml(getCustomerPenNumber(sub) || '-')}</td>
                <td>${escapeHtml(getUserDisplayName(sub.uploadedBy || ''))}</td>
                <td>${escapeHtml(getUserDisplayName(sub.auditDuplicateRejectedBy || sub.latestRejectedBy || ''))}</td>
                <td>${escapeHtml(formatDate(sub.auditDuplicateRejectedAt || sub.latestRejectedAt || sub.updatedAt))}</td>
                <td>${escapeHtml(getAuditDuplicateCorrectionLabel(sub))}</td>
                <td>${escapeHtml(getAuditDuplicateReturnedToLabel(sub))}</td>
                <td>${escapeHtml(sub.auditDuplicateRejectionReason || sub.latestRejectionReason || sub.comment || '-')}</td>
                <td>
                    <div class="audit-duplicate-actions">
                        <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
                        <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationTrack('${sub.id}')"><i class="fas fa-route"></i> Track</button>
                    </div>
                </td>
            </tr>
        `).join('')
        : `<tr><td colspan="10" class="no-data">${emptyMessage}</td></tr>`;
    syncAuditReconciliationViewVisibility();
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
    return userDisplayNamesByEmail.get(normalized) || email;
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
    if (normalized === 'audit_pending') return 'Audit Review';
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

function isReachedPfaLifecycle(sub = {}) {
    const status = String(sub.status || '').toLowerCase();
    return (
        status === 'sent_to_pfa' ||
        status === 'rsa_submitted' ||
        status === 'paid' ||
        status === 'cleared' ||
        sub.finalSubmitted === true ||
        sub.rsaSubmitted === true ||
        !!sub.finalSubmittedAt ||
        !!sub.rsaSubmittedAt ||
        !!sub.paidAt ||
        !!sub.clearedAt
    );
}

function isCurrentSentToPfaStatus(sub = {}) {
    return String(sub.status || '').toLowerCase() === 'sent_to_pfa';
}

function isClearedLifecycle(sub = {}) {
    return String(sub.status || '').toLowerCase() === 'cleared' || !!sub.clearedAt || !!sub.auditClearedAt;
}

function getSubmissionReportUserEmail(sub = {}) {
    return normalizeEmail(sub.uploadedBy || sub.auditCommissionSubmittedBy || sub.createdBy || sub.submittedBy || '');
}

function getSubmissionReportAgentName(sub = {}) {
    return String(
        sub.agentName ||
        sub.agentFullName ||
        sub.agent?.fullName ||
        sub.customerDetails?.agentName ||
        'No Agent'
    ).trim() || 'No Agent';
}

function createEmptyAuditUserReportRow(emailKey = '', user = null) {
    const email = normalizeEmail(emailKey || user?.email || '');
    return {
        emailKey: email || '__unknown__',
        email: email || '-',
        userName: String(user?.fullName || user?.name || (email ? getUserDisplayName(email) : 'Unknown User') || 'Unknown User').trim(),
        sentToPfaCount: 0,
        sentToPfaAmount: 0,
        payableCommissionCount: 0,
        payableCommissionAmount: 0,
        clearedCount: 0,
        clearedAmount: 0,
        totalCommissionCount: 0,
        totalCommissionAmount: 0,
        agentNames: new Set()
    };
}

function getAuditUserReportSentAmount(sub = {}) {
    return getSubmissionFinancials(sub).commission;
}

function getAuditUserReportClearedAmount(sub = {}) {
    return getAuditClearedCommissionAmount(sub);
}

function isPayableCommissionLifecycle(sub = {}) {
    return String(sub.status || '').toLowerCase() === 'paid';
}

function getAuditUserReportPayableAmount(sub = {}) {
    return getSubmissionFinancials(sub).commission;
}

function getAuditUserReportRows() {
    const rowsByUser = new Map();

    allUsers.forEach((user) => {
        const role = String(user.role || '').trim().toLowerCase();
        if (role && role !== 'uploader') return;
        const email = normalizeEmail(user.email || '');
        if (!email) return;
        rowsByUser.set(email, createEmptyAuditUserReportRow(email, user));
    });

    allSubmissions.forEach((sub) => {
        const email = getSubmissionReportUserEmail(sub);
        const key = email || '__unknown__';
        const row = rowsByUser.get(key) || createEmptyAuditUserReportRow(email);
        if (!row.userName || row.userName === 'Unknown User') {
            row.userName = email ? getUserDisplayName(email) : 'Unknown User';
        }
        if (isCurrentSentToPfaStatus(sub)) {
            row.sentToPfaCount += 1;
            row.sentToPfaAmount += getAuditUserReportSentAmount(sub);
            row.agentNames.add(getSubmissionReportAgentName(sub));
        }
        if (isPayableCommissionLifecycle(sub)) {
            row.payableCommissionCount += 1;
            row.payableCommissionAmount += getAuditUserReportPayableAmount(sub);
            row.agentNames.add(getSubmissionReportAgentName(sub));
        }
        if (isClearedLifecycle(sub)) {
            row.clearedCount += 1;
            row.clearedAmount += getAuditUserReportClearedAmount(sub);
            row.agentNames.add(getSubmissionReportAgentName(sub));
        }
        row.totalCommissionCount = row.sentToPfaCount + row.payableCommissionCount + row.clearedCount;
        row.totalCommissionAmount = row.sentToPfaAmount + row.payableCommissionAmount + row.clearedAmount;
        rowsByUser.set(key, row);
    });

    return Array.from(rowsByUser.values())
        .map((row) => ({
            ...row,
            agentCount: row.agentNames.size
        }))
        .sort((a, b) => b.totalCommissionAmount - a.totalCommissionAmount || b.payableCommissionAmount - a.payableCommissionAmount || b.sentToPfaAmount - a.sentToPfaAmount || b.clearedAmount - a.clearedAmount || b.sentToPfaCount - a.sentToPfaCount || a.userName.localeCompare(b.userName));
}

function getAuditAgentSummaryRowsForUser(emailKey = '') {
    const normalizedEmail = normalizeEmail(emailKey);
    const key = normalizedEmail || '__unknown__';
    const rowsByAgent = new Map();

    allSubmissions.forEach((sub) => {
        const userKey = getSubmissionReportUserEmail(sub) || '__unknown__';
        if (userKey !== key) return;
        const reachedPfa = isCurrentSentToPfaStatus(sub);
        const payable = isPayableCommissionLifecycle(sub);
        const cleared = isClearedLifecycle(sub);
        if (!reachedPfa && !payable && !cleared) return;

        const agentName = getSubmissionReportAgentName(sub);
        const agentKey = String(sub.agentId || '').trim() || agentName.toLowerCase();
        const row = rowsByAgent.get(agentKey) || {
            agentName,
            sentToPfaCount: 0,
            sentToPfaAmount: 0,
            payableCommissionCount: 0,
            payableCommissionAmount: 0,
            clearedCount: 0,
            clearedAmount: 0,
            totalCommissionCount: 0,
            totalCommissionAmount: 0
        };
        if (reachedPfa) {
            row.sentToPfaCount += 1;
            row.sentToPfaAmount += getAuditUserReportSentAmount(sub);
        }
        if (payable) {
            row.payableCommissionCount += 1;
            row.payableCommissionAmount += getAuditUserReportPayableAmount(sub);
        }
        if (cleared) {
            row.clearedCount += 1;
            row.clearedAmount += getAuditUserReportClearedAmount(sub);
        }
        row.totalCommissionCount = row.sentToPfaCount + row.payableCommissionCount + row.clearedCount;
        row.totalCommissionAmount = row.sentToPfaAmount + row.payableCommissionAmount + row.clearedAmount;
        rowsByAgent.set(agentKey, row);
    });

    return Array.from(rowsByAgent.values())
        .sort((a, b) => b.totalCommissionAmount - a.totalCommissionAmount || b.payableCommissionAmount - a.payableCommissionAmount || b.sentToPfaAmount - a.sentToPfaAmount || b.clearedAmount - a.clearedAmount || b.sentToPfaCount - a.sentToPfaCount || a.agentName.localeCompare(b.agentName));
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
    const resolvedScope = normalizeAuditPaidScope(scope);
    const currentEmail = normalizeEmail(currentUser?.email);
    return allSubmissions
        .filter((sub) => String(sub.status || '').toLowerCase() === 'paid')
        .filter((sub) => {
            if (resolvedScope === 'all') return true;
            const approvedBy = getAuditApprovalEmail(sub);
            if (!currentEmail) return resolvedScope !== 'mine';
            return resolvedScope === 'mine' ? approvedBy === currentEmail : approvedBy !== currentEmail;
        })
        .sort((a, b) => getStageTimestampMillis(b.paidAt || getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(a.paidAt || getSubmissionCurrentStageEntryAt(a)));
}

function getAuditClearedRows() {
    return allSubmissions
        .filter((sub) => String(sub.status || '').toLowerCase() === 'cleared')
        .sort((a, b) => getStageTimestampMillis(b.clearedAt || getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(a.clearedAt || getSubmissionCurrentStageEntryAt(a)));
}

function getAuditClearedCommissionAmount(sub = {}) {
    const storedAmount = parseMoney(sub.auditCommissionAmount || sub.commissionAmount || sub.commissionPaidAmount || 0);
    if (storedAmount > 0) return storedAmount;
    return getSubmissionFinancials(sub).commission;
}

function getAuditRejectedRows() {
    const currentEmail = normalizeEmail(currentUser?.email);
    return allSubmissions
        .filter((sub) => String(sub.auditCommissionStatus || '').toLowerCase() === 'rejected')
        .filter((sub) => normalizeEmail(sub.auditCommissionRejectedBy || '') === currentEmail)
        .sort((a, b) => getStageTimestampMillis(b.auditCommissionRejectedAt) - getStageTimestampMillis(a.auditCommissionRejectedAt));
}

function normalizeAuditRejectedScope(value) {
    return value === 'frozen' ? 'frozen' : 'rejected';
}

function getAuditRejectedRowsForScope(scope = currentAuditRejectedScope) {
    const resolvedScope = normalizeAuditRejectedScope(scope);
    return getAuditRejectedRows().filter((sub) => resolvedScope === 'frozen' ? sub.auditFrozen === true : sub.auditFrozen !== true);
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

function renderAuditReconciliation() {
    const selectedFile = getSelectedAuditReconciliationFile();
    if (auditReconciliationFileMeta) {
        auditReconciliationFileMeta.textContent = auditReconciliationFileName
            ? `Selected file: ${auditReconciliationFileName}`
            : (selectedFile ? `Selected file: ${selectedFile.name}` : 'No file selected.');
    }
    if (auditReconciliationRunBtn) auditReconciliationRunBtn.disabled = !selectedFile;
    if (auditReconciliationClearBtn) auditReconciliationClearBtn.disabled = !selectedFile && !auditReconciliationResult && !auditDuplicateScanResult;
    if (auditReconciliationExportBtn) auditReconciliationExportBtn.disabled = !auditReconciliationResult && !auditDuplicateScanResult;

    if (!auditReconciliationResult) {
        if (auditReconciliationSummary) {
            auditReconciliationSummary.style.display = 'none';
            auditReconciliationSummary.innerHTML = '';
        }
        if (auditReconciliationMatchedWrap) auditReconciliationMatchedWrap.style.display = 'none';
        if (auditReconciliationMatchedBody) auditReconciliationMatchedBody.innerHTML = '';
        if (auditReconciliationUnmatchedWrap) auditReconciliationUnmatchedWrap.style.display = 'none';
        if (auditReconciliationUnmatchedBody) auditReconciliationUnmatchedBody.innerHTML = '';
        syncAuditReconciliationViewVisibility();
        return;
    }

    const matchedCount = auditReconciliationResult.matchedRows.length;
    const unmatchedCount = auditReconciliationResult.unmatchedRows.length;
    if (auditReconciliationSummary) {
        auditReconciliationSummary.style.display = 'grid';
        auditReconciliationSummary.innerHTML = [
            { label: 'Uploaded Rows', value: auditReconciliationResult.totalRows },
            { label: 'Found', value: matchedCount },
            { label: 'Exact Balance', value: auditReconciliationResult.exactCount },
            { label: 'Balance Differs', value: auditReconciliationResult.balanceDifferenceCount },
            { label: 'Not Found', value: unmatchedCount },
            { label: 'System Records', value: auditReconciliationResult.systemCount }
        ].map((chip) => `
            <div class="audit-reconciliation-chip">
                <span>${escapeHtml(chip.label)}</span>
                <strong>${String(chip.value)}</strong>
            </div>
        `).join('');
    }

    if (auditReconciliationMatchedWrap) auditReconciliationMatchedWrap.style.display = matchedCount ? 'block' : 'none';
    if (auditReconciliationMatchedBody) {
        auditReconciliationMatchedBody.innerHTML = matchedCount
            ? auditReconciliationResult.matchedRows.map((row) => {
                const statusClass = row.balanceMatches ? 'audit-recon-status exact' : 'audit-recon-status partial';
                const statusText = row.balanceMatches ? `Found (${row.matchMethod})` : `Found, balance differs (${row.matchMethod})`;
                return `
                    <tr>
                        <td>${escapeHtml(row.rowNumber)}</td>
                        <td>${escapeHtml(row.customerName || '-')}</td>
                        <td>${escapeHtml(row.accountNumber || '-')}</td>
                        <td>${row.hasRsaBalance ? formatCurrency(row.normalizedRsaBalance) : '-'}</td>
                        <td><strong>${escapeHtml(row.systemName || '-')}</strong></td>
                        <td>${escapeHtml(row.uploaderName || '-')}</td>
                        <td>${escapeHtml(row.systemAccountNumber || '-')}</td>
                        <td>${formatCurrency(row.systemRsaBalance)}</td>
                        <td><span class="${statusClass}">${escapeHtml(statusText)}</span></td>
                        <td><button type="button" class="action-btn" onclick="window.openMonitoringApplicationDetails('${row.submissionId}')"><i class="fas fa-eye"></i> View</button></td>
                    </tr>
                `;
            }).join('')
            : '<tr><td colspan="10" class="no-data">No found records</td></tr>';
    }

    if (auditReconciliationUnmatchedWrap) auditReconciliationUnmatchedWrap.style.display = unmatchedCount ? 'block' : 'none';
    if (auditReconciliationUnmatchedBody) {
        auditReconciliationUnmatchedBody.innerHTML = unmatchedCount
            ? auditReconciliationResult.unmatchedRows.map((row) => `
                <tr>
                    <td>${escapeHtml(row.rowNumber)}</td>
                    <td>${escapeHtml(row.customerName || '-')}</td>
                    <td>${escapeHtml(row.accountNumber || '-')}</td>
                    <td>${row.hasRsaBalance ? formatCurrency(row.normalizedRsaBalance) : '-'}</td>
                    <td>${escapeHtml(row.reason || 'Not found')}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="5" class="no-data">No missing records</td></tr>';
    }
    syncAuditReconciliationViewVisibility();
}

async function downloadAuditReconciliationTemplate() {
    if (!window.ExcelJS) throw new Error('Excel library is not available right now.');
    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Audit Reconciliation');
    sheet.columns = [
        { header: 'Customer Name', key: 'customerName', width: 34 },
        { header: 'Account Number', key: 'accountNumber', width: 18 },
        { header: 'RSA Balance', key: 'rsaBalance', width: 18 }
    ];
    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DCEBFA' } };
    sheet.getColumn(2).numFmt = '@';
    sheet.addRow({ customerName: 'Sample Customer', accountNumber: '0123456789', rsaBalance: 1000000 });
    sheet.addRow({ customerName: '', accountNumber: '', rsaBalance: '' });
    const buffer = await workbook.xlsx.writeBuffer();
    saveBlob('audit-reconciliation-template.xlsx', new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
}

async function exportAuditReconciliationResult() {
    if (!auditReconciliationResult && !auditDuplicateScanResult) throw new Error('Run reconciliation or duplicate scan first.');
    if (!window.ExcelJS) throw new Error('Excel library is not available right now.');
    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Audit Dashboard';
    workbook.created = new Date();

    if (auditReconciliationResult) {
        const foundSheet = workbook.addWorksheet('Found Records');
        foundSheet.columns = [
            { header: 'Upload Row', key: 'rowNumber', width: 12 },
            { header: 'Uploaded Name', key: 'customerName', width: 30 },
            { header: 'Uploaded Account', key: 'accountNumber', width: 18 },
            { header: 'Uploaded RSA Balance', key: 'uploadedRsaBalance', width: 20 },
            { header: 'System Name', key: 'systemName', width: 30 },
            { header: 'Uploader', key: 'uploaderName', width: 28 },
            { header: 'System Account', key: 'systemAccountNumber', width: 18 },
            { header: 'System RSA Balance', key: 'systemRsaBalance', width: 20 },
            { header: 'System Status', key: 'systemStatus', width: 16 },
            { header: 'Match Status', key: 'matchStatus', width: 26 },
            { header: 'Application ID', key: 'submissionId', width: 28 }
        ];
        auditReconciliationResult.matchedRows.forEach((row) => foundSheet.addRow({
            rowNumber: row.rowNumber,
            customerName: row.customerName,
            accountNumber: row.accountNumber,
            uploadedRsaBalance: row.hasRsaBalance ? row.normalizedRsaBalance : '',
            systemName: row.systemName,
            uploaderName: row.uploaderName,
            systemAccountNumber: row.systemAccountNumber,
            systemRsaBalance: row.systemRsaBalance,
            systemStatus: row.systemStatus,
            matchStatus: row.balanceMatches ? `Found (${row.matchMethod})` : `Found, balance differs (${row.matchMethod})`,
            submissionId: row.submissionId
        }));
        foundSheet.getRow(1).font = { bold: true };

        const notFoundSheet = workbook.addWorksheet('Not Found');
        notFoundSheet.columns = [
            { header: 'Upload Row', key: 'rowNumber', width: 12 },
            { header: 'Name', key: 'customerName', width: 30 },
            { header: 'Account Number', key: 'accountNumber', width: 18 },
            { header: 'RSA Balance', key: 'rsaBalance', width: 20 },
            { header: 'Result', key: 'reason', width: 48 }
        ];
        auditReconciliationResult.unmatchedRows.forEach((row) => notFoundSheet.addRow({
            rowNumber: row.rowNumber,
            customerName: row.customerName,
            accountNumber: row.accountNumber,
            rsaBalance: row.hasRsaBalance ? row.normalizedRsaBalance : '',
            reason: row.reason || 'Not found'
        }));
        notFoundSheet.getRow(1).font = { bold: true };
    }

    if (auditDuplicateScanResult) {
        const duplicateSheet = workbook.addWorksheet('Duplicate Applications');
        duplicateSheet.columns = [
            { header: 'Group', key: 'group', width: 10 },
            { header: 'Duplicate Signal', key: 'signal', width: 42 },
            { header: 'Signal Strength', key: 'strength', width: 16 },
            { header: 'Customer', key: 'customerName', width: 30 },
            { header: 'Account Number', key: 'accountNumber', width: 18 },
            { header: 'PEN', key: 'penNo', width: 22 },
            { header: 'Status', key: 'status', width: 18 },
            { header: 'Uploader', key: 'uploader', width: 28 },
            { header: 'Uploaded At', key: 'uploadedAt', width: 24 },
            { header: 'Application ID', key: 'submissionId', width: 28 }
        ];
        (auditDuplicateScanResult.groups || []).forEach((group, groupIndex) => {
            (group.rows || []).forEach((sub) => duplicateSheet.addRow({
                group: groupIndex + 1,
                signal: group.key || group.signal || '',
                strength: group.strength || '',
                customerName: sub.customerName || sub?.customerDetails?.name || '',
                accountNumber: getCustomerAccountNumber(sub),
                penNo: getCustomerPenNumber(sub),
                status: statusLabel(sub.status || ''),
                uploader: getUserDisplayName(sub.uploadedBy || ''),
                uploadedAt: formatDate(getSubmissionOriginalUploadAt(sub)),
                submissionId: sub.id || ''
            }));
        });
        duplicateSheet.getRow(1).font = { bold: true };
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveBlob(`audit-reconciliation-result-${Date.now()}.xlsx`, new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
}

function renderOverview() {
    const awaitingAuditCount = getAuditSentToPfaRows().length;
    const paidRows = getAuditPaidRows('all');
    const paidCount = paidRows.length;
    const clearedRows = getAuditClearedRows();
    const clearedCount = clearedRows.length;
    const commissionPayableAmount = paidRows.reduce((sum, sub) => sum + getSubmissionFinancials(sub).commission, 0);
    const clearedAmount = clearedRows.reduce((sum, sub) => sum + getAuditClearedCommissionAmount(sub), 0);

    setCountBadge('overviewUsersCount', awaitingAuditCount);
    setCountBadge('overviewSentCount', paidCount);
    setCountBadge('overviewCommissionPayableAmount', formatCurrency(commissionPayableAmount));
    setCountBadge('overviewPaidClearedCount', clearedCount);
    setCountBadge('auditClearedTotalApplications', clearedCount);
    setCountBadge('auditClearedTotalAmount', formatCurrency(clearedAmount));
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
            : mode === 'paid'
                ? `<div class="audit-paid-actions">
                       <button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
                       <button class="action-btn audit-clear-btn" onclick="window.clearAuditPayment('${sub.id}')"><i class="fas fa-circle-check"></i> Clear</button>
                   </div>`
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
        const isFrozen = sub.auditFrozen === true;
        const freezeButton = isFrozen
            ? `<button class="action-btn audit-unfreeze-btn" onclick="window.toggleAuditApplicationFreeze('${sub.id}', false)" title="Allow the uploader to act on this application"><i class="fas fa-lock-open"></i> Unfreeze</button>`
            : `<button class="action-btn audit-freeze-btn" onclick="window.toggleAuditApplicationFreeze('${sub.id}', true)" title="Prevent the uploader from changing this application"><i class="fas fa-lock"></i> Freeze</button>`;
        const actionCell = `<div class="audit-rejected-actions">
            <button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
            ${freezeButton}
            ${isFrozen ? '<span class="audit-frozen-pill"><i class="fas fa-snowflake"></i> Frozen</span>' : ''}
        </div>`;

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

function renderCommissionSummaryMetric(count = 0, amount = 0) {
    return `
        <div class="commission-summary-metric">
            <strong>${formatCurrency(amount || 0)}</strong>
            <span>${escapeHtml(String(count || 0))} application${Number(count || 0) === 1 ? '' : 's'}</span>
        </div>
    `;
}

function renderAuditUserReport() {
    const rows = getAuditUserReportRows();
    const activeRows = rows.filter((row) => row.sentToPfaCount > 0 || row.payableCommissionCount > 0 || row.clearedCount > 0);
    const totalSentAmount = rows.reduce((sum, row) => sum + row.sentToPfaAmount, 0);
    const totalClearedAmount = rows.reduce((sum, row) => sum + row.clearedAmount, 0);

    if (auditUserReportUserCount) auditUserReportUserCount.textContent = String(activeRows.length);
    if (auditUserReportSentCount) auditUserReportSentCount.textContent = formatCurrency(totalSentAmount);
    if (auditUserReportClearedCount) auditUserReportClearedCount.textContent = formatCurrency(totalClearedAmount);

    if (!auditUserReportTableBody) return;
    auditUserReportTableBody.innerHTML = rows.length
        ? rows.map((row) => {
            const encodedEmail = encodeURIComponent(row.emailKey || '');
            return `
                <tr>
                    <td><strong>${escapeHtml(row.userName || 'Unknown User')}</strong></td>
                    <td>${escapeHtml(row.email || '-')}</td>
                    <td>${renderCommissionSummaryMetric(row.sentToPfaCount, row.sentToPfaAmount)}</td>
                    <td>${renderCommissionSummaryMetric(row.payableCommissionCount, row.payableCommissionAmount)}</td>
                    <td>${renderCommissionSummaryMetric(row.clearedCount, row.clearedAmount)}</td>
                    <td>${renderCommissionSummaryMetric(row.totalCommissionCount, row.totalCommissionAmount)}</td>
                    <td>${escapeHtml(String(row.agentCount || 0))}</td>
                    <td><button type="button" class="action-btn" onclick="window.openAuditUserAgentSummary('${encodedEmail}')"><i class="fas fa-eye"></i> View</button></td>
                </tr>
            `;
        }).join('')
        : '<tr><td colspan="8" class="no-data">No user report data available</td></tr>';
}

function openAuditUserAgentSummary(encodedEmailKey = '') {
    const emailKey = decodeURIComponent(String(encodedEmailKey || ''));
    const rows = getAuditAgentSummaryRowsForUser(emailKey);
    const userRow = getAuditUserReportRows().find((row) => row.emailKey === emailKey) || createEmptyAuditUserReportRow(emailKey);
    const totalSentCount = rows.reduce((sum, row) => sum + row.sentToPfaCount, 0);
    const totalSentAmount = rows.reduce((sum, row) => sum + row.sentToPfaAmount, 0);
    const totalPayableCount = rows.reduce((sum, row) => sum + row.payableCommissionCount, 0);
    const totalPayableAmount = rows.reduce((sum, row) => sum + row.payableCommissionAmount, 0);
    const totalClearedCount = rows.reduce((sum, row) => sum + row.clearedCount, 0);
    const totalClearedAmount = rows.reduce((sum, row) => sum + row.clearedAmount, 0);
    const totalCommissionCount = rows.reduce((sum, row) => sum + row.totalCommissionCount, 0);
    const totalCommissionAmount = rows.reduce((sum, row) => sum + row.totalCommissionAmount, 0);

    if (auditUserAgentSummaryTitle) {
        auditUserAgentSummaryTitle.innerHTML = '<i class="fas fa-chart-column"></i> Agent Summary';
    }
    if (auditUserAgentSummaryName) {
        auditUserAgentSummaryName.textContent = userRow.userName || 'User';
    }
    if (auditUserAgentSummaryMeta) {
        auditUserAgentSummaryMeta.textContent = `${userRow.email || '-'} | ${rows.length} agent${rows.length === 1 ? '' : 's'} covered`;
    }
    if (auditUserAgentSummaryCards) {
        auditUserAgentSummaryCards.innerHTML = [
            { label: 'Sent to PFA', amount: formatCurrency(totalSentAmount), count: totalSentCount, tone: 'sent', icon: 'fa-paper-plane' },
            { label: 'Payable Commission', amount: formatCurrency(totalPayableAmount), count: totalPayableCount, tone: 'payable', icon: 'fa-money-check-dollar' },
            { label: 'Cleared', amount: formatCurrency(totalClearedAmount), count: totalClearedCount, tone: 'cleared', icon: 'fa-circle-check' },
            { label: 'Total Commission', amount: formatCurrency(totalCommissionAmount), count: totalCommissionCount, tone: 'total', icon: 'fa-coins' }
        ].map((chip) => `
            <div class="commission-summary-card ${escapeHtml(chip.tone)}">
                <div class="commission-summary-card-icon"><i class="fas ${escapeHtml(chip.icon)}"></i></div>
                <span>${escapeHtml(chip.label)}</span>
                <strong>${escapeHtml(chip.amount)}</strong>
                <small>${escapeHtml(String(chip.count))} application${Number(chip.count) === 1 ? '' : 's'}</small>
            </div>
        `).join('');
    }
    if (auditUserAgentSummaryBody) {
        auditUserAgentSummaryBody.innerHTML = rows.length
            ? rows.map((row) => `
                <tr>
                    <td><strong>${escapeHtml(row.agentName || 'No Agent')}</strong></td>
                    <td>${renderCommissionSummaryMetric(row.sentToPfaCount, row.sentToPfaAmount)}</td>
                    <td>${renderCommissionSummaryMetric(row.payableCommissionCount, row.payableCommissionAmount)}</td>
                    <td>${renderCommissionSummaryMetric(row.clearedCount, row.clearedAmount)}</td>
                    <td>${renderCommissionSummaryMetric(row.totalCommissionCount, row.totalCommissionAmount)}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="5" class="no-data">No agent summary available for this user</td></tr>';
    }
    auditUserAgentSummaryModal?.classList.add('active');
}

function closeAuditUserAgentSummaryModal() {
    auditUserAgentSummaryModal?.classList.remove('active');
}

function renderAuditWorkflowTabs() {
    renderAuditWorkflowBadges();
    document.querySelectorAll('[data-audit-paid-scope]').forEach((button) => {
        button.classList.toggle('active', button.dataset.auditPaidScope === currentAuditPaidScope);
    });
    document.querySelectorAll('[data-audit-rejected-scope]').forEach((button) => {
        button.classList.toggle('active', button.dataset.auditRejectedScope === currentAuditRejectedScope);
    });
    renderAuditMoneyRows(auditOverviewPendingTableBody, getAuditSentToPfaRows(), 'sent');
    renderAuditMoneyRows(auditSentToPfaTableBody, getAuditSentToPfaRows(), 'sent');
    renderAuditMoneyRows(auditPaidTableBody, getAuditPaidRows(currentAuditPaidScope), 'paid');
    renderAuditMoneyRows(auditClearedTableBody, getAuditClearedRows(), 'cleared');
    renderAuditRejectedRows(auditRejectedTableBody, getAuditRejectedRowsForScope(currentAuditRejectedScope));
    renderAuditPaidReconciliation();
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
    const agentSummaryMap = new Map();
    const details = records
        .map((sub) => {
            const { twentyFive, commission } = getSubmissionFinancials(sub);
            const dateMs = getDateMs(sub);
            const attended = attendedResolver(sub);
            const statusText = statusResolver(sub);
            const agentName = String(sub?.agentName || '').trim() || 'No Agent';
            const agentAccountNumber = String(sub?.agentAccountNumber || '').trim() || '-';
            const agentBank = String(sub?.agentAccountBank || '').trim() || '-';
            const uploaderName = getUserDisplayName(sub?.uploadedBy);
            const detailRow = {
                id: sub.id,
                receivedAtMs: dateMs,
                receivedDate: formatDateOnly(dateMs),
                customerName: String(sub?.customerName || 'Unknown').trim() || 'Unknown',
                customerAccountNumber: getCustomerAccountNumber(sub),
                agentName,
                agentAccountNumber,
                agentBank,
                uploaderName,
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

            const agentSummaryKey = String(sub?.agentId || '').trim() || `${agentName.toLowerCase()}::${normalizeEmail(sub?.uploadedBy)}::${agentAccountNumber}::${agentBank.toLowerCase()}`;
            const existingAgentSummary = agentSummaryMap.get(agentSummaryKey) || {
                agentName,
                uploaderName,
                agentAccountNumber,
                agentBank,
                totalCommissionPayable: 0
            };
            if ((!existingAgentSummary.uploaderName || existingAgentSummary.uploaderName === '-') && uploaderName && uploaderName !== '-') {
                existingAgentSummary.uploaderName = uploaderName;
            }
            if ((!existingAgentSummary.agentAccountNumber || existingAgentSummary.agentAccountNumber === '-') && agentAccountNumber && agentAccountNumber !== '-') {
                existingAgentSummary.agentAccountNumber = agentAccountNumber;
            }
            if ((!existingAgentSummary.agentBank || existingAgentSummary.agentBank === '-') && agentBank && agentBank !== '-') {
                existingAgentSummary.agentBank = agentBank;
            }
            existingAgentSummary.totalCommissionPayable += commission;
            agentSummaryMap.set(agentSummaryKey, existingAgentSummary);

            return detailRow;
        })
        .sort((a, b) => b.receivedAtMs - a.receivedAtMs);

    const summaryRows = Array.from(summaryMap.values()).sort((a, b) => String(a.key).localeCompare(String(b.key)));
    const agentSummaryRows = Array.from(agentSummaryMap.values())
        .sort((a, b) => b.totalCommissionPayable - a.totalCommissionPayable || a.agentName.localeCompare(b.agentName) || String(a.uploaderName || '').localeCompare(String(b.uploaderName || '')));
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
        metaText: `${title} generated from ${records.length} record(s). Includes agent commission summary and application breakdown for the selected ${dateLabel.toLowerCase()} range.`,
        summaryRows,
        agentSummaryRows,
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

    const paidScope = normalizeAuditPaidScope(request?.scope || currentAuditPaidScope);
    const records = getAuditPaidRows(paidScope).filter((sub) => {
        const dateMs = getTimestampMillis(sub.paidAt);
        return dateMs >= startMs && dateMs <= endMs;
    });
    return buildPaymentStageReport(records, {
        title: `Paid Report - ${getAuditPaidScopeLabel(paidScope)}`,
        exportKey: `paid-report-${paidScope}-${startDate}-to-${endDate}`,
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
    const totalCommissionPayable = report.details.reduce((total, row) => total + Number(row.commission || 0), 0);
    if (paymentReportSummaryChips) {
        paymentReportSummaryChips.innerHTML = [
            { label: 'Agents Covered', value: report.agentSummaryRows.length, bg: '#dbeafe', color: '#1d4ed8' },
            { label: 'Applications', value: report.details.length, bg: '#e2e8f0', color: '#334155' },
            { label: 'Total Commission Payable', value: formatCurrency(totalCommissionPayable), bg: '#dcfce7', color: '#166534' }
        ].map((item) => `<span style="display:inline-flex;align-items:center;gap:6px;padding:8px 12px;border-radius:999px;background:${item.bg};color:${item.color};font-size:12px;font-weight:700;"><span>${escapeHtml(item.label)}</span><span>${escapeHtml(String(item.value))}</span></span>`).join('');
    }
    if (paymentReportSummaryBody) {
        paymentReportSummaryBody.innerHTML = report.agentSummaryRows.length
            ? report.agentSummaryRows.map((row) => `<tr><td>${escapeHtml(row.uploaderName || '-')}</td><td>${escapeHtml(row.agentName)}</td><td>${escapeHtml(row.agentAccountNumber || '-')}</td><td>${escapeHtml(row.agentBank || '-')}</td><td>${formatCurrency(row.totalCommissionPayable)}</td></tr>`).join('')
            : '<tr><td colspan="5" class="no-data">No commission breakdown available for this date range</td></tr>';
    }
    if (paymentReportDetailsBody) {
        paymentReportDetailsBody.innerHTML = report.details.length
            ? report.details.map((row) => `
                <tr>
                    <td><strong>${escapeHtml(row.customerName)}</strong></td>
                    <td>${escapeHtml(row.customerAccountNumber || '-')}</td>
                    <td>${escapeHtml(row.uploaderName || '-')}</td>
                    <td>${formatCurrency(row.twentyFive)}</td>
                    <td>${formatCurrency(row.commission)}</td>
                    <td>${escapeHtml(row.rateLabel || '-')}</td>
                    <td>${escapeHtml(row.agentName)}</td>
                    <td>${escapeHtml(row.agentAccountNumber || '-')}</td>
                    <td>${escapeHtml(row.agentBank || '-')}</td>
                    <td>${escapeHtml(row.statusLabel)}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="10" class="no-data">No application breakdown available for this date range</td></tr>';
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
    workbook.creator = 'CMBank RSA Audit Dashboard';
    workbook.created = new Date();

    const summarySheet = workbook.addWorksheet('Agent Summary');
    summarySheet.columns = [
        { header: 'Uploader Name', key: 'uploaderName', width: 28 },
        { header: 'Agent Name', key: 'agentName', width: 30 },
        { header: 'Agent Account Number', key: 'agentAccountNumber', width: 22 },
        { header: 'Agent Bank', key: 'agentBank', width: 24 },
        { header: 'Total Commission Payable', key: 'totalCommissionPayable', width: 26, style: { numFmt: '#,##0.00' } }
    ];
    summarySheet.addRow({ uploaderName: 'Range', agentName: `${report.rangeStart} to ${report.rangeEnd}` });
    summarySheet.addRow({});
    (report.agentSummaryRows || []).forEach((row) => summarySheet.addRow(row));
    summarySheet.addRow({});
    summarySheet.addRow({
        uploaderName: 'Grand Total',
        totalCommissionPayable: (report.agentSummaryRows || []).reduce((total, row) => total + Number(row.totalCommissionPayable || 0), 0)
    });
    summarySheet.getRow(1).font = { bold: true };
    summarySheet.getRow(summarySheet.rowCount).font = { bold: true };

    const detailsSheet = workbook.addWorksheet('Application Breakdown');
    detailsSheet.columns = [
        { header: 'Customer Name', key: 'customerName', width: 28 },
        { header: 'Customer Account Number', key: 'customerAccountNumber', width: 22 },
        { header: 'Uploader Name', key: 'uploaderName', width: 24 },
        { header: '25% RSA Balance', key: 'twentyFive', width: 18, style: { numFmt: '#,##0.00' } },
        { header: 'Commission', key: 'commission', width: 18, style: { numFmt: '#,##0.00' } },
        { header: 'Rate', key: 'rateLabel', width: 12 },
        { header: 'Agent Name', key: 'agentName', width: 24 },
        { header: 'Agent Account Number', key: 'agentAccountNumber', width: 22 },
        { header: 'Agent Bank', key: 'agentBank', width: 22 },
        { header: 'Status', key: 'statusLabel', width: 16 },
    ];
    report.details.forEach((row) => detailsSheet.addRow(row));
    detailsSheet.getColumn(2).numFmt = '@';
    detailsSheet.getColumn(8).numFmt = '@';
    detailsSheet.getRow(1).font = { bold: true };

    const buffer = await workbook.xlsx.writeBuffer();
    saveBlob(`${report.exportKey || 'payment-report'}.xlsx`, new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
}

async function exportPaymentReportPdf(report) {
    if (!report || !window.jspdf?.jsPDF) throw new Error('PDF library not available.');
    const pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const totalCommissionPayable = report.details.reduce((total, row) => total + Number(row.commission || 0), 0);
    pdf.setFontSize(16);
    pdf.text(report.title || 'Payment Report', 40, 36);
    pdf.setFontSize(10);
    pdf.text(report.metaText || '', 40, 54);
    pdf.text(`Agents: ${report.agentSummaryRows.length}   Applications: ${report.details.length}   Total Commission Payable: ${formatCurrency(totalCommissionPayable)}`, 40, 70);
    pdf.autoTable({
        startY: 88,
        head: [['Uploader Name', 'Agent Name', 'Agent Account Number', 'Agent Bank', 'Total Commission Payable']],
        body: report.agentSummaryRows.map((row) => [
            row.uploaderName || '-',
            row.agentName,
            row.agentAccountNumber || '-',
            row.agentBank || '-',
            Number(row.totalCommissionPayable || 0).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        ]),
        styles: { fontSize: 9 },
        headStyles: { fillColor: [15, 59, 103] }
    });
    pdf.autoTable({
        startY: pdf.lastAutoTable.finalY + 18,
        head: [['Customer', 'Customer Account', 'Uploader', '25% RSA Balance', 'Commission', 'Rate', 'Agent', 'Agent Account', 'Agent Bank', 'Status']],
        body: report.details.map((row) => [
            row.customerName,
            row.customerAccountNumber || '-',
            row.uploaderName,
            Number(row.twentyFive || 0).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
            Number(row.commission || 0).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
            row.rateLabel,
            row.agentName,
            row.agentAccountNumber || '-',
            row.agentBank || '-',
            row.statusLabel,
        ]),
        styles: { fontSize: 7, cellPadding: 2.5, overflow: 'linebreak', valign: 'middle' },
        headStyles: { fillColor: [15, 118, 110] },
        showHead: 'everyPage',
        margin: { left: 18, right: 18 },
        tableWidth: 'auto',
        columnStyles: {
            0: { cellWidth: 96 },
            1: { cellWidth: 72 },
            2: { cellWidth: 82 },
            3: { cellWidth: 60, halign: 'right' },
            4: { cellWidth: 60, halign: 'right' },
            5: { cellWidth: 36, halign: 'center' },
            6: { cellWidth: 80 },
            7: { cellWidth: 72 },
            8: { cellWidth: 68 },
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
    renderAuditWorkflowBadges();

    if (currentTab === 'overview') {
        renderOverview();
        renderAuditMoneyRows(auditOverviewPendingTableBody, getAuditSentToPfaRows(), 'sent');
        return;
    }
    if (currentTab === 'sent-to-pfa') {
        renderAuditMoneyRows(auditSentToPfaTableBody, getAuditSentToPfaRows(), 'sent');
        return;
    }
    if (currentTab === 'paid') {
        document.querySelectorAll('[data-audit-paid-scope]').forEach((button) => {
            button.classList.toggle('active', button.dataset.auditPaidScope === currentAuditPaidScope);
        });
        renderAuditMoneyRows(auditPaidTableBody, getAuditPaidRows(currentAuditPaidScope), 'paid');
        renderAuditPaidReconciliation();
        return;
    }
    if (currentTab === 'cleared') {
        renderAuditMoneyRows(auditClearedTableBody, getAuditClearedRows(), 'cleared');
        return;
    }
    if (currentTab === 'rejected') {
        document.querySelectorAll('[data-audit-rejected-scope]').forEach((button) => {
            button.classList.toggle('active', button.dataset.auditRejectedScope === currentAuditRejectedScope);
        });
        renderAuditRejectedRows(auditRejectedTableBody, getAuditRejectedRowsForScope(currentAuditRejectedScope));
        return;
    }
    if (currentTab === 'reconciliation') {
        renderAuditReconciliation();
        renderAuditDuplicateScan();
        renderAuditDuplicateIgnoredTable();
        renderAuditDuplicateRejectedTable();
        renderAuditPaidReconciliation();
        return;
    }
    if (currentTab === 'user-report') {
        renderAuditUserReport();
    }
}

function scheduleCurrentTabRender() {
    clearTimeout(auditRenderTimer);
    auditRenderTimer = setTimeout(renderCurrentTab, 40);
}

function switchTab(tabId) {
    tabId = AUDIT_DASHBOARD_TABS.includes(tabId) ? tabId : 'overview';
    currentTab = tabId;
    rememberAuditTab(tabId);
    ensureDataForTab(tabId);
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');
    renderCurrentTab();

    const titles = {
        overview: 'Audit Overview',
        'sent-to-pfa': 'Pending Request',
        paid: 'Paid',
        cleared: 'Cleared',
        rejected: 'Rejected',
        reconciliation: 'Reconciliation',
        'user-report': 'Commission Summary',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Audit';
}

function ensureDataForTab(tabId) {
    const dataTabs = ['overview', 'sent-to-pfa', 'paid', 'cleared', 'rejected', 'reconciliation', 'user-report'];
    if (dataTabs.includes(tabId)) loadUsers();
    if (dataTabs.includes(tabId)) loadSubmissions({ full: tabId === 'reconciliation' || tabId === 'user-report' });
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
        ['Uploaded At', formatDate(getSubmissionOriginalUploadAt(sub))],
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

function openApplicationTrackModal(submissionId) {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const rows = [
        ['Application ID', sub.id || '-'],
        ['Customer Name', sub.customerName || '-'],
        ['Current Status', statusLabel(sub.status || '-')],
        ['Current Stage', getApplicationStage(sub)],
        ['Original Upload', formatDate(getSubmissionOriginalUploadAt(sub))],
        ['Current Stage Since', formatDate(getSubmissionCurrentStageEntryAt(sub))],
        ['Uploader', sub.uploadedBy || '-'],
        ['Reviewer', sub.assignedTo || sub.reviewedBy || '-'],
        ['RSA Officer', sub.assignedToRSA || '-'],
        ['Payment Officer', sub.assignedToPayment || '-'],
        ['Paid At', formatDate(sub.paidAt)],
        ['Cleared At', formatDate(sub.clearedAt || sub.auditClearedAt)],
        ['Audit Status', statusLabel(sub.auditCommissionStatus || '-')]
    ];

    const title = document.getElementById('applicationDetailsTitle');
    const body = document.getElementById('applicationDetailsBody');
    if (title) title.textContent = `Application Track - ${sub.customerName || 'Customer'}`;
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

function getSubmissionUniqueKeyDocId(type, value) {
    return encodeURIComponent(`${String(type || '').trim().toLowerCase()}:${String(value || '').trim()}`);
}

function getSubmissionUniqueKeyRefs(sub = {}) {
    const keys = [
        ['account_number', normalizeAccountNumber(getCustomerAccountNumber(sub))],
        ['nin', getCustomerNin(sub)],
        ['pen', normalizePenNumber(getCustomerPenNumber(sub))],
        ['phone', getCustomerPhoneNumber(sub)],
        ['customer_name', normalizeCustomerName(sub.customerName || sub?.customerDetails?.name || '')]
    ].filter(([, value]) => String(value || '').trim());
    return keys.map(([type, value]) => doc(db, 'submissionUniqueKeys', getSubmissionUniqueKeyDocId(type, value)));
}

function getAuditDuplicateCustomerName(sub = {}) {
    return String(sub?.customerName || sub?.customerDetails?.name || 'Unknown application').trim();
}

function formatAuditDuplicateNameList(names = []) {
    const uniqueNames = Array.from(new Set(names.map((name) => String(name || '').trim()).filter(Boolean)));
    if (!uniqueNames.length) return '';
    if (uniqueNames.length === 1) return uniqueNames[0];
    if (uniqueNames.length === 2) return `${uniqueNames[0]} and ${uniqueNames[1]}`;
    return `${uniqueNames.slice(0, -1).join(', ')}, and ${uniqueNames[uniqueNames.length - 1]}`;
}

function getAuditDuplicateGroupBySubmissionId(submissionId = '') {
    const id = String(submissionId || '').trim();
    if (!id) return null;
    if (!auditDuplicateScanResult) auditDuplicateScanResult = buildAuditDuplicateScanResult();
    return (auditDuplicateScanResult?.groups || []).find((group) =>
        (group.rows || []).some((row) => String(row.id || '') === id)
    ) || null;
}

function getAuditDuplicateGroupByIndex(groupIndex) {
    const index = Number(groupIndex);
    if (!Number.isInteger(index) || index < 0) return null;
    if (!auditDuplicateScanResult) auditDuplicateScanResult = buildAuditDuplicateScanResult();
    return auditDuplicateScanResult?.groups?.[index] || null;
}

function getAuditDuplicateNinSignal(group = {}) {
    const signal = (group.signals || []).find((item) => String(item.signal || '').toLowerCase() === 'nin');
    if (signal?.key) return String(signal.key || '').trim();
    const rows = group.rows || [];
    const ninValues = rows.map(getCustomerNin).filter(Boolean);
    return ninValues.length && ninValues.every((value) => value === ninValues[0]) ? ninValues[0] : '';
}

function getAuditDuplicateSharedNames(sub = {}, group = {}) {
    const submissionId = String(sub?.id || '').trim();
    return (group.rows || [])
        .filter((row) => String(row.id || '').trim() !== submissionId)
        .map(getAuditDuplicateCustomerName)
        .filter(Boolean);
}

function buildAuditDuplicateRejectionReason(sub = {}, group = null, enteredReason = '') {
    const duplicateGroup = group || getAuditDuplicateGroupBySubmissionId(sub.id) || { rows: [sub], signals: [] };
    const sharedNamesText = formatAuditDuplicateNameList(getAuditDuplicateSharedNames(sub, duplicateGroup));
    const ninSignal = getAuditDuplicateNinSignal(duplicateGroup);
    const defaultSpecificReason = sharedNamesText && ninSignal
        ? `Duplicate application found. Please correct the duplicated NIN shared with ${sharedNamesText}.`
        : sharedNamesText
            ? `Duplicate application found. Please correct the duplicated customer/account details shared with ${sharedNamesText}.`
            : 'Duplicate application found. Please correct the duplicated customer/account details and resubmit.';
    const reason = String(enteredReason || '').trim();
    if (!reason || reason === defaultSpecificReason) return defaultSpecificReason;
    if (!sharedNamesText) return reason;
    if (reason.toLowerCase().includes(sharedNamesText.toLowerCase())) return reason;
    return ninSignal
        ? `${reason} Duplicated NIN shared with ${sharedNamesText}.`
        : `${reason} Shared with ${sharedNamesText}.`;
}

function showAuditDuplicateRejectModal({ submission = {}, group = null } = {}) {
    return new Promise((resolve) => {
        const rows = group?.rows || [submission];
        const isGroup = !submission?.id && rows.length > 1;
        const primarySubmission = submission?.id ? submission : rows[0] || {};
        const defaultReason = buildAuditDuplicateRejectionReason(primarySubmission, group);
        const title = isGroup ? 'Reject Duplicate Group' : 'Reject Duplicate Application';
        const targetText = isGroup
            ? `${rows.length} duplicate applications`
            : `<strong>${escapeHtml(getAuditDuplicateCustomerName(primarySubmission) || 'this application')}</strong>`;
        const modal = document.createElement('div');
        modal.className = 'modal active audit-action-modal';
        modal.innerHTML = `
            <div class="modal-content audit-action-card reject">
                <div class="audit-action-icon">
                    <i class="fas fa-ban"></i>
                </div>
                <h2>${escapeHtml(title)}</h2>
                <p>Send ${targetText} back to the uploader for correction.</p>
                <textarea id="auditDuplicateRejectReasonInput" rows="4" placeholder="Reason for rejection">${escapeHtml(defaultReason)}</textarea>
                <div class="audit-action-actions">
                    <button type="button" class="cancel-btn" data-audit-duplicate-reject="cancel">Cancel</button>
                    <button type="button" class="submit-btn danger" data-audit-duplicate-reject="confirm">
                        <i class="fas fa-paper-plane"></i> Reject
                    </button>
                </div>
            </div>
        `;
        const close = (value) => {
            modal.remove();
            resolve(value);
        };
        modal.addEventListener('click', (event) => {
            if (event.target === modal || event.target.closest('[data-audit-duplicate-reject="cancel"]')) {
                close('');
                return;
            }
            if (!event.target.closest('[data-audit-duplicate-reject="confirm"]')) return;
            const reasonInput = modal.querySelector('#auditDuplicateRejectReasonInput');
            const reason = String(reasonInput?.value || '').trim();
            if (!reason) {
                reasonInput?.classList.add('invalid');
                return;
            }
            close(reason);
        });
        document.body.appendChild(modal);
        setTimeout(() => modal.querySelector('#auditDuplicateRejectReasonInput')?.focus(), 0);
    });
}

async function rejectAuditDuplicateSubmissions(submissions = [], enteredReason = '', group = null) {
    const rows = Array.from(new Map(
        submissions
            .filter((sub) => sub?.id)
            .map((sub) => [sub.id, sub])
    ).values());
    if (!rows.length) return 0;
    const rejectedBy = currentUser?.email || '';
    const batch = writeBatch(db);
    const auditEntries = [];

    rows.forEach((sub) => {
        const submissionId = String(sub.id || '').trim();
        if (!submissionId) return;
        const reason = buildAuditDuplicateRejectionReason(sub, group, enteredReason);
        const customerName = getAuditDuplicateCustomerName(sub) || 'this application';
        const previousStatus = String(sub.status || '').trim();
        const previousStage = getApplicationStage(sub);
        batch.update(doc(db, 'submissions', submissionId), {
            status: 'rejected',
            comment: reason,
            rejectionHistory: arrayUnion({
                reason,
                rejectedAt: new Date().toISOString(),
                rejectedBy,
                source: 'audit_duplicate_scan'
            }),
            latestRejectionReason: reason,
            latestRejectedBy: rejectedBy,
            latestRejectedAt: serverTimestamp(),
            previousRejectionReason: reason,
            previousRejectedBy: rejectedBy,
            previousRejectedAt: serverTimestamp(),
            resubmittedAfterRejection: false,
            latestRejectedStage: 'audit',
            auditDuplicateRejected: true,
            auditDuplicateRejectedBy: rejectedBy,
            auditDuplicateRejectedAt: serverTimestamp(),
            auditDuplicateRejectionReason: reason,
            auditDuplicatePreviousStatus: previousStatus,
            auditDuplicatePreviousStage: previousStage,
            updatedAt: serverTimestamp()
        });
        auditEntries.push({
            action: rows.length > 1 ? 'audit_duplicate_group_application_rejected' : 'audit_duplicate_application_rejected',
            submissionId,
            customerName,
            accountNumber: getCustomerAccountNumber(sub),
            nin: getCustomerNin(sub),
            penNo: getCustomerPenNumber(sub),
            reason,
            previousStatus,
            previousStage,
            duplicateGroupSize: rows.length,
            performedBy: rejectedBy,
            timestamp: serverTimestamp()
        });
    });

    await batch.commit();
    await Promise.all(auditEntries.map((entry) => addDoc(collection(db, 'audit'), entry).catch(() => {})));

    await Promise.all(rows.map((sub) => {
        const submissionId = String(sub.id || '').trim();
        const reason = buildAuditDuplicateRejectionReason(sub, group, enteredReason);
        const customerName = getAuditDuplicateCustomerName(sub) || 'this application';
        return notifyUserPushEvent({
            currentUser,
            recipientEmail: String(sub.uploadedBy || '').trim(),
            eventType: 'audit_duplicate_application_rejected',
            title: 'Application Needs Correction',
            body: `${customerName} was rejected by Audit because it appears duplicated.`,
            clickUrl: '/dashboard.html#rejected',
            meta: {
                submissionId,
                customerName,
                reason,
                rejectedBy
            }
        }).catch(() => {});
    }));

    rows.forEach((sub) => {
        const localSub = allSubmissions.find((item) => item.id === sub.id);
        if (!localSub) return;
        const reason = buildAuditDuplicateRejectionReason(sub, group, enteredReason);
        localSub.status = 'rejected';
        localSub.comment = reason;
        localSub.latestRejectionReason = reason;
        localSub.latestRejectedBy = rejectedBy;
        localSub.previousRejectionReason = reason;
        localSub.previousRejectedBy = rejectedBy;
        localSub.resubmittedAfterRejection = false;
        localSub.latestRejectedStage = 'audit';
        localSub.auditDuplicateRejected = true;
        localSub.auditDuplicateRejectedAt = new Date().toISOString();
        localSub.auditDuplicateRejectedBy = rejectedBy;
        localSub.auditDuplicateRejectionReason = reason;
        localSub.auditDuplicatePreviousStatus = String(sub.status || '').trim();
        localSub.auditDuplicatePreviousStage = getApplicationStage(sub);
    });

    return rows.length;
}

function showAuditConfirmModal({ title = 'Confirm Action', message = '', confirmLabel = 'Confirm', icon = 'fa-check', danger = false } = {}) {
    return new Promise((resolve) => {
        const modal = document.createElement('div');
        modal.className = 'modal active audit-action-modal';
        modal.innerHTML = `
            <div class="modal-content audit-action-card ${danger ? 'reject' : 'accept'}">
                <div class="audit-action-icon">
                    <i class="fas ${escapeHtml(icon)}"></i>
                </div>
                <h2>${escapeHtml(title)}</h2>
                <p>${escapeHtml(message)}</p>
                <div class="audit-action-actions">
                    <button type="button" class="cancel-btn" data-audit-confirm="cancel">Cancel</button>
                    <button type="button" class="submit-btn ${danger ? 'danger' : ''}" data-audit-confirm="confirm">
                        <i class="fas ${escapeHtml(icon)}"></i> ${escapeHtml(confirmLabel)}
                    </button>
                </div>
            </div>
        `;
        const close = (value) => {
            modal.remove();
            resolve(value);
        };
        modal.addEventListener('click', (event) => {
            if (event.target === modal || event.target.closest('[data-audit-confirm="cancel"]')) {
                close(false);
                return;
            }
            if (event.target.closest('[data-audit-confirm="confirm"]')) close(true);
        });
        document.body.appendChild(modal);
    });
}

async function rejectAuditDuplicateApplication(submissionId) {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const group = getAuditDuplicateGroupBySubmissionId(submissionId);
    const reason = await showAuditDuplicateRejectModal({ submission: sub, group });
    if (!reason) return;

    try {
        await rejectAuditDuplicateSubmissions([sub], reason, group);

        auditDuplicateScanResult = buildAuditDuplicateScanResult();
        showAuditReconciliationView('duplicates');
        showNotification('Duplicate application rejected for uploader correction.', 'success');
    } catch (error) {
        showNotification(`Failed to reject duplicate application: ${error.message || error}`, 'error');
    }
}

async function rejectAuditDuplicateGroup(groupIndex) {
    const group = getAuditDuplicateGroupByIndex(groupIndex);
    const rows = (group?.rows || []).filter((sub) => sub?.id);
    if (!rows.length) {
        showNotification('Duplicate group not found', 'warning');
        return;
    }

    const reason = await showAuditDuplicateRejectModal({ group });
    if (!reason) return;

    try {
        const defaultGroupReason = buildAuditDuplicateRejectionReason(rows[0], group);
        const enteredReason = reason === defaultGroupReason ? '' : reason;
        const rejectedCount = await rejectAuditDuplicateSubmissions(rows, enteredReason, group);
        auditDuplicateScanResult = buildAuditDuplicateScanResult();
        showAuditReconciliationView('duplicates');
        showNotification(`${rejectedCount} duplicate application(s) rejected for uploader correction.`, 'success');
    } catch (error) {
        showNotification(`Failed to reject duplicate group: ${error.message || error}`, 'error');
    }
}

async function ignoreAuditDuplicateApplication(submissionId) {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const customerName = sub.customerName || sub?.customerDetails?.name || 'this application';
    const confirmed = await showAuditConfirmModal({
        title: 'Ignore Duplicate',
        message: `Ignore "${customerName}" in duplicate checks? It will no longer appear in Find Duplicates unless this flag is changed in the database.`,
        confirmLabel: 'Ignore',
        icon: 'fa-eye-slash'
    });
    if (!confirmed) return;

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            auditDuplicateIgnored: true,
            auditDuplicateIgnoredAt: serverTimestamp(),
            auditDuplicateIgnoredBy: currentUser?.email || '',
            updatedAt: serverTimestamp()
        });

        await addDoc(collection(db, 'audit'), {
            action: 'audit_duplicate_application_ignored',
            submissionId,
            customerName,
            accountNumber: getCustomerAccountNumber(sub),
            penNo: getCustomerPenNumber(sub),
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        }).catch(() => {});

        const localSub = allSubmissions.find((item) => item.id === submissionId);
        if (localSub) {
            localSub.auditDuplicateIgnored = true;
            localSub.auditDuplicateIgnoredAt = new Date().toISOString();
            localSub.auditDuplicateIgnoredBy = currentUser?.email || '';
        }
        auditDuplicateScanResult = buildAuditDuplicateScanResult();
        showAuditReconciliationView('duplicates');
        showNotification('Application ignored for future duplicate scans.', 'success');
    } catch (error) {
        showNotification(`Failed to ignore duplicate application: ${error.message || error}`, 'error');
    }
}

function showAuditDuplicateDeleteModal(submission = {}) {
    return new Promise((resolve) => {
        const modal = document.createElement('div');
        modal.className = 'modal active audit-action-modal';
        modal.innerHTML = `
            <div class="modal-content audit-action-card reject">
                <div class="audit-action-icon">
                    <i class="fas fa-trash"></i>
                </div>
                <h2>Delete Duplicate Application</h2>
                <p>Move <strong>${escapeHtml(submission.customerName || 'this application')}</strong> to the uploader's Deleted records.</p>
                <textarea id="auditDuplicateDeleteReasonInput" rows="4" placeholder="Reason for delete">Duplicate application deleted by Audit.</textarea>
                <div class="audit-action-actions">
                    <button type="button" class="cancel-btn" data-audit-duplicate-delete="cancel">Cancel</button>
                    <button type="button" class="submit-btn danger" data-audit-duplicate-delete="confirm">
                        <i class="fas fa-trash"></i> Delete
                    </button>
                </div>
            </div>
        `;
        const close = (value) => {
            modal.remove();
            resolve(value);
        };
        modal.addEventListener('click', (event) => {
            if (event.target === modal || event.target.closest('[data-audit-duplicate-delete="cancel"]')) {
                close('');
                return;
            }
            if (!event.target.closest('[data-audit-duplicate-delete="confirm"]')) return;
            const reasonInput = modal.querySelector('#auditDuplicateDeleteReasonInput');
            const reason = String(reasonInput?.value || '').trim();
            if (!reason) {
                reasonInput?.classList.add('invalid');
                return;
            }
            close(reason);
        });
        document.body.appendChild(modal);
        setTimeout(() => modal.querySelector('#auditDuplicateDeleteReasonInput')?.focus(), 0);
    });
}

async function getOwnedSubmissionUniqueKeyRefs(sub = {}) {
    const submissionId = String(sub?.id || '').trim();
    if (!submissionId) return [];
    const refs = getSubmissionUniqueKeyRefs(sub);
    const ownedRefs = [];
    await Promise.all(refs.map(async (keyRef) => {
        try {
            const snap = await getDoc(keyRef);
            if (snap.exists() && String(snap.data()?.submissionId || '').trim() === submissionId) {
                ownedRefs.push(keyRef);
            }
        } catch (_) {}
    }));
    return ownedRefs;
}

async function deleteAuditDuplicateApplication(submissionId) {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const customerName = sub.customerName || sub?.customerDetails?.name || 'this application';
    const reason = await showAuditDuplicateDeleteModal(sub);
    if (!reason) return;

    try {
        const batch = writeBatch(db);
        batch.update(doc(db, 'submissions', submissionId), {
            status: 'deleted',
            deletedAt: serverTimestamp(),
            deletedBy: currentUser?.email || '',
            deletedReason: reason,
            auditDuplicateDeleted: true,
            auditDuplicateDeletedAt: serverTimestamp(),
            auditDuplicateDeletedBy: currentUser?.email || '',
            auditDuplicateDeleteReason: reason,
            updatedAt: serverTimestamp()
        });
        const ownedKeyRefs = await getOwnedSubmissionUniqueKeyRefs(sub);
        ownedKeyRefs.forEach((keyRef) => batch.delete(keyRef));
        await batch.commit();

        await addDoc(collection(db, 'audit'), {
            action: 'audit_duplicate_application_deleted',
            submissionId,
            customerName,
            accountNumber: getCustomerAccountNumber(sub),
            penNo: getCustomerPenNumber(sub),
            reason,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        }).catch(() => {});

        const localSub = allSubmissions.find((item) => item.id === submissionId);
        if (localSub) {
            localSub.status = 'deleted';
            localSub.deletedAt = new Date().toISOString();
            localSub.deletedBy = currentUser?.email || '';
            localSub.deletedReason = reason;
            localSub.auditDuplicateDeleted = true;
            localSub.auditDuplicateDeletedAt = localSub.deletedAt;
            localSub.auditDuplicateDeletedBy = currentUser?.email || '';
            localSub.auditDuplicateDeleteReason = reason;
        }
        auditDuplicateScanResult = buildAuditDuplicateScanResult();
        showAuditReconciliationView('duplicates');
        showNotification('Duplicate application moved to uploader Deleted records.', 'success');
    } catch (error) {
        showNotification(`Failed to delete duplicate application: ${error.message || error}`, 'error');
    }
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
        const isBulkClear = mode === 'clear-multiple';
        const isClear = mode === 'clear' || isBulkClear;
        const title = isReject ? 'Reject Payment Request' : isBulkClear ? 'Clear Multiple Payments' : isClear ? 'Clear Payment' : 'Approve Payment Request';
        const message = isReject
            ? 'Enter a reason for rejecting'
            : isClear
                ? 'Confirm that payment should be cleared for'
                : 'Confirm that commission payment should be marked as paid for';
        const modal = document.createElement('div');
        modal.className = 'modal active audit-action-modal';
        modal.innerHTML = `
            <div class="modal-content audit-action-card ${isReject ? 'reject' : 'accept'}">
                <div class="audit-action-icon">
                    <i class="fas ${isReject ? 'fa-xmark' : isClear ? 'fa-circle-check' : 'fa-check'}"></i>
                </div>
                <h2>${title}</h2>
                <p>${message} <strong>${escapeHtml(submission.customerName || 'this application')}</strong>.</p>
                ${isBulkClear ? '' : getAuditCorrectionSummaryHtml(submission)}
                ${isReject ? '<textarea id="auditRejectReasonInput" rows="4" placeholder="Enter rejection reason"></textarea>' : ''}
                <div class="audit-action-actions">
                    <button type="button" class="cancel-btn" data-audit-action="cancel">Cancel</button>
                    <button type="button" class="submit-btn ${isReject ? 'danger' : ''}" data-audit-action="confirm">
                        <i class="fas ${isReject ? 'fa-paper-plane' : isClear ? 'fa-circle-check' : 'fa-check'}"></i> ${isReject ? 'Reject' : isClear ? 'Clear' : 'Approve'}
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

function showAuditSuccessModal({ title = 'Success', message = '', detail = '' } = {}) {
    return new Promise((resolve) => {
        const modal = document.createElement('div');
        modal.className = 'modal active audit-action-modal';
        modal.innerHTML = `
            <div class="modal-content audit-action-card accept">
                <div class="audit-action-icon">
                    <i class="fas fa-circle-check"></i>
                </div>
                <h2>${escapeHtml(title)}</h2>
                <p>${escapeHtml(message)}</p>
                ${detail ? `<p style="font-size:13px;color:#64748b;">${escapeHtml(detail)}</p>` : ''}
                <div class="audit-action-actions">
                    <button type="button" class="submit-btn" data-audit-success="ok">
                        <i class="fas fa-check"></i> OK
                    </button>
                </div>
            </div>
        `;
        const close = () => {
            modal.remove();
            resolve(true);
        };
        modal.addEventListener('click', (event) => {
            if (event.target === modal || event.target.closest('[data-audit-success="ok"]')) close();
        });
        document.body.appendChild(modal);
    });
}

window.openMonitoringApplicationDetails = openApplicationDetailsModal;
window.openMonitoringApplicationTrack = openApplicationTrackModal;
window.ignoreAuditDuplicateApplication = ignoreAuditDuplicateApplication;
window.rejectAuditDuplicateApplication = rejectAuditDuplicateApplication;
window.rejectAuditDuplicateGroup = rejectAuditDuplicateGroup;
window.deleteAuditDuplicateApplication = deleteAuditDuplicateApplication;

function getAuditPaymentClearPayload(sub = {}) {
    const { commission, twentyFive, rsaBalance } = getSubmissionFinancials(sub);
    return {
        updates: {
            status: 'cleared',
            clearedAt: serverTimestamp(),
            clearedBy: currentUser?.email || '',
            auditClearedAt: serverTimestamp(),
            auditClearedBy: currentUser?.email || '',
            auditCommissionAmount: commission,
            auditRsaBalance: rsaBalance,
            auditRsaTwentyFivePercent: twentyFive,
            updatedAt: serverTimestamp()
        },
        auditLog: {
            action: 'audit_payment_cleared',
            submissionId: String(sub?.id || '').trim(),
            customerName: sub.customerName || '',
            uploadedBy: sub.uploadedBy || '',
            commissionAmount: commission,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        },
        notification: {
            currentUser,
            recipientEmail: String(sub.auditCommissionSubmittedBy || sub.paymentMadeBy || sub.uploadedBy || '').trim(),
            eventType: 'audit_payment_cleared',
            title: 'Payment Cleared',
            body: `${sub.customerName || 'Your application'} has been cleared by Audit.`,
            clickUrl: '/dashboard.html',
            meta: {
                submissionId: String(sub?.id || '').trim(),
                customerName: sub.customerName || '',
                clearedBy: currentUser?.email || '',
                commissionAmount: commission
            }
        }
    };
}

function queueAuditPaymentClearSideEffects(records = []) {
    if (!records.length) return;
    setTimeout(() => {
        records.forEach(({ sub, payload }) => {
            addDoc(collection(db, 'audit'), payload.auditLog).catch(() => {});
            notifyUserPushEvent(payload.notification).catch(() => {});
        });
    }, 0);
}

function markAuditPaymentRecordsClearedLocally(records = []) {
    records.forEach(({ sub }) => {
        const submissionId = String(sub?.id || '').trim();
        const localSub = allSubmissions.find((item) => item.id === submissionId);
        if (!localSub) return;
        localSub.status = 'cleared';
        localSub.clearedBy = currentUser?.email || '';
        localSub.auditClearedBy = currentUser?.email || '';
        localSub.auditCommissionStatus = 'cleared';
    });
}

async function clearAuditPaymentRecordsBulk(submissions = [], { onProgress = null } = {}) {
    const clearableRecords = [];
    let skippedCount = 0;

    submissions.forEach((sub) => {
        const submissionId = String(sub?.id || '').trim();
        if (!submissionId || String(sub.status || '').toLowerCase() !== 'paid') {
            skippedCount += 1;
            return;
        }
        clearableRecords.push({ sub, payload: getAuditPaymentClearPayload(sub) });
    });

    let clearedCount = 0;
    let failedCount = 0;
    const clearedRecords = [];

    for (let index = 0; index < clearableRecords.length; index += AUDIT_BULK_CLEAR_BATCH_SIZE) {
        const chunk = clearableRecords.slice(index, index + AUDIT_BULK_CLEAR_BATCH_SIZE);
        const batch = writeBatch(db);
        chunk.forEach(({ sub, payload }) => {
            batch.update(doc(db, 'submissions', String(sub.id)), payload.updates);
        });

        try {
            await batch.commit();
            clearedCount += chunk.length;
            clearedRecords.push(...chunk);
            markAuditPaymentRecordsClearedLocally(chunk);
            if (typeof onProgress === 'function') onProgress(clearedCount, clearableRecords.length);
        } catch (_) {
            failedCount += chunk.length;
        }
    }

    queueAuditPaymentClearSideEffects(clearedRecords);
    return { clearedCount, skippedCount, failedCount };
}

async function clearAuditPaymentRecord(sub = {}) {
    const submissionId = String(sub?.id || '').trim();
    if (!submissionId) throw new Error('Application not found.');
    if (String(sub.status || '').toLowerCase() !== 'paid') throw new Error('Only paid applications can be cleared.');

    const payload = getAuditPaymentClearPayload(sub);
    await updateDoc(doc(db, 'submissions', submissionId), payload.updates);

    await addDoc(collection(db, 'audit'), payload.auditLog).catch(() => {});

    await notifyUserPushEvent(payload.notification).catch(() => {});
}

window.clearAuditPayment = async (submissionId) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }
    if (String(sub.status || '').toLowerCase() !== 'paid') {
        showNotification('Only paid applications can be cleared.', 'warning');
        return;
    }

    const confirmed = await showAuditActionModal({ mode: 'clear', submission: sub });
    if (!confirmed) return;

    try {
        await clearAuditPaymentRecord(sub);
        showNotification('Payment cleared successfully.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to clear payment.', 'error');
    }
};

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

window.toggleAuditApplicationFreeze = async (submissionId, shouldFreeze) => {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }
    if (String(sub.auditCommissionStatus || '').toLowerCase() !== 'rejected') {
        showNotification('Only rejected Audit applications can be frozen or unfrozen.', 'warning');
        return;
    }

    const action = shouldFreeze ? 'freeze' : 'unfreeze';
    const confirmed = await showAuditConfirmModal({
        title: shouldFreeze ? 'Freeze Application' : 'Unfreeze Application',
        message: shouldFreeze
            ? `Freeze ${sub.customerName || 'this application'}? The uploader will not be able to resubmit, dissolve, chat, or make other changes until Audit unfreezes it.`
            : `Unfreeze ${sub.customerName || 'this application'}? The uploader will be able to act on it again.`,
        confirmLabel: shouldFreeze ? 'Freeze' : 'Unfreeze',
        icon: shouldFreeze ? 'fa-snowflake' : 'fa-unlock',
        danger: shouldFreeze
    });
    if (!confirmed) return;

    try {
        const actorEmail = currentUser?.email || '';
        await updateDoc(doc(db, 'submissions', submissionId), shouldFreeze ? {
            auditFrozen: true,
            auditFrozenAt: serverTimestamp(),
            auditFrozenBy: actorEmail,
            auditUnfrozenAt: null,
            auditUnfrozenBy: '',
            updatedAt: serverTimestamp()
        } : {
            auditFrozen: false,
            auditUnfrozenAt: serverTimestamp(),
            auditUnfrozenBy: actorEmail,
            updatedAt: serverTimestamp()
        });

        await addDoc(collection(db, 'audit'), {
            action: shouldFreeze ? 'audit_application_frozen' : 'audit_application_unfrozen',
            submissionId,
            customerName: sub.customerName || '',
            uploadedBy: sub.uploadedBy || '',
            performedBy: actorEmail,
            timestamp: serverTimestamp()
        }).catch(() => {});

        await notifyUserPushEvent({
            currentUser,
            recipientEmail: String(sub.auditCommissionSubmittedBy || sub.paymentMadeBy || sub.uploadedBy || '').trim(),
            eventType: shouldFreeze ? 'audit_application_frozen' : 'audit_application_unfrozen',
            title: shouldFreeze ? 'Application Frozen by Audit' : 'Application Unfrozen by Audit',
            body: shouldFreeze
                ? `${sub.customerName || 'Your application'} is frozen and cannot be changed until Audit unfreezes it.`
                : `${sub.customerName || 'Your application'} has been unfrozen and can be actioned again.`,
            clickUrl: '/dashboard.html',
            meta: {
                submissionId,
                customerName: sub.customerName || '',
                performedBy: actorEmail
            }
        }).catch(() => {});

        sub.auditFrozen = shouldFreeze;
        if (shouldFreeze) {
            sub.auditFrozenBy = actorEmail;
            sub.auditUnfrozenBy = '';
        } else {
            sub.auditUnfrozenBy = actorEmail;
        }
        renderCurrentTab();
        showNotification(`Application ${shouldFreeze ? 'frozen' : 'unfrozen'} successfully.`, 'success');
    } catch (error) {
        showNotification(`Failed to ${action} application.`, 'error');
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
            currentAuditPaidScope = normalizeAuditPaidScope(button.dataset.auditPaidScope);
            if (auditPaidReconciliationSourceRows.length) recomputeAuditPaidReconciliationResult();
            renderCurrentTab();
        });
    });
    document.querySelectorAll('[data-audit-rejected-scope]').forEach((button) => {
        button.addEventListener('click', () => {
            currentAuditRejectedScope = normalizeAuditRejectedScope(button.dataset.auditRejectedScope);
            renderCurrentTab();
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
    document.getElementById('forceRefreshBtn')?.addEventListener('click', (event) => {
        event.preventDefault();
        forceHardRefresh();
    });
    document.getElementById('forceRefreshBtnMobile')?.addEventListener('click', (event) => {
        event.preventDefault();
        forceHardRefresh();
    });
    exportAuditPendingReportBtn?.addEventListener('click', () => {
        pendingPaymentReportRequest = { kind: 'sent-to-pfa' };
        openPaymentReportRangeModal();
    });
    exportAuditPaidReportBtn?.addEventListener('click', () => {
        pendingPaymentReportRequest = { kind: 'paid', scope: currentAuditPaidScope };
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
    document.querySelectorAll('[data-audit-reconciliation-view]').forEach((button) => {
        button.addEventListener('click', () => showAuditReconciliationView(button.dataset.auditReconciliationView || 'excel'));
    });
    auditDuplicateHistoryFilter?.addEventListener('change', () => {
        currentAuditDuplicateHistoryFilter = normalizeAuditDuplicateHistoryFilter(auditDuplicateHistoryFilter.value);
        renderAuditDuplicateRejectedTable();
    });
    auditReconciliationTemplateBtn?.addEventListener('click', async () => {
        const originalHtml = auditReconciliationTemplateBtn.innerHTML;
        auditReconciliationTemplateBtn.disabled = true;
        auditReconciliationTemplateBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Template';
        try {
            await downloadAuditReconciliationTemplate();
            showNotification('Template downloaded successfully.', 'success');
        } catch (error) {
            showNotification(error?.message || 'Failed to download template', 'error');
        } finally {
            auditReconciliationTemplateBtn.disabled = false;
            auditReconciliationTemplateBtn.innerHTML = originalHtml;
        }
    });
    auditReconciliationSelectBtn?.addEventListener('click', () => {
        if (auditReconciliationFileInput) auditReconciliationFileInput.value = '';
        auditReconciliationFileInput?.click();
    });
    auditReconciliationFileInput?.addEventListener('change', () => {
        const selectedFile = getSelectedAuditReconciliationFile();
        auditReconciliationFileName = selectedFile?.name || '';
        auditReconciliationSourceRows = [];
        auditReconciliationResult = null;
        auditReconciliationActiveView = 'excel';
        renderAuditReconciliation();
        if (selectedFile) setTimeout(() => showNotification('File uploaded. Click Run Check to start reconciliation.', 'success'), 0);
    });
    auditReconciliationRunBtn?.addEventListener('click', async () => {
        const file = getSelectedAuditReconciliationFile();
        if (!file) {
            showNotification('Select a file first.', 'warning');
            return;
        }
        const originalHtml = auditReconciliationRunBtn.innerHTML;
        auditReconciliationRunBtn.disabled = true;
        auditReconciliationRunBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Checking...';
        try {
            auditReconciliationFileName = file.name;
            auditReconciliationSourceRows = await parseAuditReconciliationFile(file);
            auditReconciliationResult = computeAuditReconciliationResult(auditReconciliationSourceRows);
            auditReconciliationActiveView = 'excel';
            renderAuditReconciliation();
            showNotification(`Reconciliation complete. ${auditReconciliationResult.matchedRows.length} row(s) found.`, 'success');
        } catch (error) {
            auditReconciliationSourceRows = [];
            auditReconciliationResult = null;
            renderAuditReconciliation();
            showNotification(error?.message || 'Failed to run reconciliation', 'error');
        } finally {
            auditReconciliationRunBtn.disabled = !getSelectedAuditReconciliationFile();
            auditReconciliationRunBtn.innerHTML = originalHtml;
        }
    });
    auditReconciliationExportBtn?.addEventListener('click', async () => {
        const originalHtml = auditReconciliationExportBtn.innerHTML;
        auditReconciliationExportBtn.disabled = true;
        auditReconciliationExportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Exporting...';
        try {
            await exportAuditReconciliationResult();
            showNotification('Reconciliation result exported.', 'success');
        } catch (error) {
            showNotification(error?.message || 'Failed to export result', 'error');
        } finally {
            auditReconciliationExportBtn.disabled = !auditReconciliationResult && !auditDuplicateScanResult;
            auditReconciliationExportBtn.innerHTML = originalHtml;
        }
    });
    auditReconciliationClearBtn?.addEventListener('click', () => {
        resetAuditReconciliationState();
        showNotification('Reconciliation cleared.', 'info');
    });
    openAuditPaidReconciliationBtn?.addEventListener('click', openAuditPaidReconciliationUploadModal);
    document.getElementById('closeAuditPaidReconciliationUploadModalBtn')?.addEventListener('click', closeAuditPaidReconciliationUploadModal);
    document.getElementById('cancelAuditPaidReconciliationUploadModalBtn')?.addEventListener('click', closeAuditPaidReconciliationUploadModal);
    document.getElementById('closeAuditPaidReconciliationResultsModalBtn')?.addEventListener('click', closeAuditPaidReconciliationResultsModal);
    document.getElementById('closeAuditPaidReconciliationResultsFooterBtn')?.addEventListener('click', closeAuditPaidReconciliationResultsModal);
    auditPaidReconciliationUploadModal?.addEventListener('click', (event) => {
        if (event.target === auditPaidReconciliationUploadModal) closeAuditPaidReconciliationUploadModal();
    });
    auditPaidReconciliationResultsModal?.addEventListener('click', (event) => {
        if (event.target === auditPaidReconciliationResultsModal) closeAuditPaidReconciliationResultsModal();
    });
    document.querySelectorAll('[data-paid-reconciliation-view]').forEach((button) => {
        button.addEventListener('click', () => showAuditPaidReconciliationResultsTab(button.dataset.paidReconciliationView || 'matched'));
    });
    auditPaidReconciliationSelectBtn?.addEventListener('click', () => {
        if (auditPaidReconciliationFileInput) auditPaidReconciliationFileInput.value = '';
        auditPaidReconciliationFileInput?.click();
    });
    auditPaidReconciliationFileInput?.addEventListener('change', () => {
        const selectedFile = getSelectedAuditPaidReconciliationFile();
        auditPaidReconciliationFileName = selectedFile?.name || '';
        auditPaidReconciliationSourceRows = [];
        auditPaidReconciliationResult = null;
        auditPaidReconciliationSelectedSubmissionIds = [];
        renderAuditPaidReconciliation();
        if (selectedFile) setTimeout(() => showNotification('File uploaded. Click Run Check to match paid applications.', 'success'), 0);
    });
    auditPaidReconciliationRunBtn?.addEventListener('click', async () => {
        const file = getSelectedAuditPaidReconciliationFile();
        if (!file) {
            showNotification('Select a file first.', 'warning');
            return;
        }
        const originalHtml = auditPaidReconciliationRunBtn.innerHTML;
        auditPaidReconciliationRunBtn.disabled = true;
        auditPaidReconciliationRunBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Checking...';
        try {
            auditPaidReconciliationFileName = file.name;
            auditPaidReconciliationSourceRows = await parseAuditReconciliationFile(file);
            auditPaidReconciliationSelectedSubmissionIds = [];
            recomputeAuditPaidReconciliationResult();
            auditPaidReconciliationSelectedSubmissionIds = Array.from(new Set(
                (auditPaidReconciliationResult?.matchedRows || [])
                    .filter((row) => isAuditPaidRowClearable(row))
                    .map((row) => row.submissionId)
                    .filter(Boolean)
            ));
            auditPaidReconciliationActiveResultsTab = (auditPaidReconciliationResult?.matchedRows || []).length ? 'matched' : 'unmatched';
            renderAuditPaidReconciliation();
            closeAuditPaidReconciliationUploadModal();
            openAuditPaidReconciliationResultsModal();
            showNotification(`Reconciliation complete. ${auditPaidReconciliationResult?.matchedRows?.length || 0} application(s) found in the system, ${auditPaidReconciliationSelectedSubmissionIds.length} selected for clearing.`, 'success');
        } catch (error) {
            auditPaidReconciliationSourceRows = [];
            auditPaidReconciliationResult = null;
            auditPaidReconciliationSelectedSubmissionIds = [];
            renderAuditPaidReconciliation();
            showNotification(error?.message || 'Failed to run paid reconciliation.', 'error');
        } finally {
            auditPaidReconciliationRunBtn.disabled = !getSelectedAuditPaidReconciliationFile();
            auditPaidReconciliationRunBtn.innerHTML = originalHtml;
        }
    });
    auditPaidReconciliationSelectAllBtn?.addEventListener('click', () => {
        auditPaidReconciliationSelectedSubmissionIds = Array.from(new Set(
            (auditPaidReconciliationResult?.matchedRows || [])
                .filter((row) => isAuditPaidRowClearable(row))
                .map((row) => row.submissionId)
                .filter(Boolean)
        ));
        renderAuditPaidReconciliation();
    });
    auditPaidReconciliationMatchedBody?.addEventListener('change', (event) => {
        const checkbox = event.target.closest('[data-paid-recon-select]');
        if (!checkbox) return;
        const submissionId = String(checkbox.dataset.paidReconSelect || '').trim();
        if (!submissionId) return;
        const selected = new Set(auditPaidReconciliationSelectedSubmissionIds);
        if (checkbox.checked) selected.add(submissionId);
        else selected.delete(submissionId);
        auditPaidReconciliationSelectedSubmissionIds = Array.from(selected);
        renderAuditPaidReconciliation();
    });
    auditPaidReconciliationClearSelectedBtn?.addEventListener('click', async () => {
        const submissionIds = Array.from(new Set(auditPaidReconciliationSelectedSubmissionIds.filter(Boolean)));
        if (!submissionIds.length) {
            showNotification('Select at least one paid application to clear.', 'warning');
            return;
        }
        const confirmed = await showAuditActionModal({
            mode: 'clear-multiple',
            submission: {
                customerName: `${submissionIds.length} selected application${submissionIds.length === 1 ? '' : 's'}`
            }
        });
        if (!confirmed) return;

        const originalHtml = auditPaidReconciliationClearSelectedBtn.innerHTML;
        auditPaidReconciliationClearSelectedBtn.disabled = true;
        auditPaidReconciliationClearSelectedBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Clearing...';
        auditPaidReconciliationSelectedSubmissionIds = [];
        renderAuditPaidReconciliation();

        let result = { clearedCount: 0, skippedCount: 0, failedCount: 0 };
        try {
            const selectedSubmissions = submissionIds.map((submissionId) => allSubmissions.find((item) => item.id === submissionId)).filter(Boolean);
            result = await clearAuditPaymentRecordsBulk(selectedSubmissions, {
                onProgress: (cleared, total) => {
                    auditPaidReconciliationClearSelectedBtn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Clearing ${cleared}/${total}...`;
                }
            });

            const parts = [];
            if (result.clearedCount) parts.push(`Cleared ${result.clearedCount} application(s)`);
            if (result.skippedCount) parts.push(`Skipped ${result.skippedCount}`);
            if (result.failedCount) parts.push(`Failed ${result.failedCount}`);
            const message = parts.length ? `${parts.join('. ')}.` : 'No payments were cleared.';
            showNotification(message, result.failedCount ? (result.clearedCount ? 'warning' : 'error') : 'success');
            if (!result.failedCount && result.clearedCount) {
                closeAuditPaidReconciliationResultsModal();
                closeAuditPaidReconciliationUploadModal();
                resetAuditPaidReconciliationState();
                renderAuditPaidReconciliation();
                renderCurrentTab();
                await showAuditSuccessModal({
                    title: 'Clearing Complete',
                    message: `${result.clearedCount} application${result.clearedCount === 1 ? '' : 's'} cleared successfully.`,
                    detail: result.skippedCount ? `${result.skippedCount} already-cleared or unavailable application${result.skippedCount === 1 ? '' : 's'} skipped.` : ''
                });
            }
        } finally {
            if (auditPaidReconciliationSourceRows.length) recomputeAuditPaidReconciliationResult();
            auditPaidReconciliationClearSelectedBtn.innerHTML = originalHtml;
            renderAuditPaidReconciliation();
        }
    });
    auditPaidReconciliationResetBtn?.addEventListener('click', () => {
        resetAuditPaidReconciliationState();
        renderAuditPaidReconciliation();
        closeAuditPaidReconciliationResultsModal();
        showNotification('Paid reconciliation cleared.', 'info');
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
    document.getElementById('closeAuditUserAgentSummaryModalBtn')?.addEventListener('click', closeAuditUserAgentSummaryModal);
    document.getElementById('closeAuditUserAgentSummaryModalFooterBtn')?.addEventListener('click', closeAuditUserAgentSummaryModal);
    window.addEventListener('click', (e) => {
        if (e.target === applicationDetailsModal) closeApplicationDetailsModal();
        if (e.target === auditUserAgentSummaryModal) closeAuditUserAgentSummaryModal();
    });
}

window.openAuditUserAgentSummary = openAuditUserAgentSummary;

function loadUsers() {
    if (usersListenerStarted) return;
    usersListenerStarted = true;
    onSnapshot(collection(db, 'users'), (snapshot) => {
        allUsers = snapshot.docs
            .map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }))
            .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));
        userDisplayNamesByEmail = new Map(allUsers
            .map((user) => [normalizeEmail(user.email), String(user.fullName || user.email || '').trim()])
            .filter(([email]) => email));
        scheduleCurrentTabRender();
    }, () => {
        showNotification('Failed to load users', 'error');
    });
}

function loadSubmissions({ full = false } = {}) {
    const mode = full ? 'full' : 'audit';
    if (submissionsListenerMode === mode) return;

    submissionListenerUnsubs.forEach((unsubscribe) => {
        try { unsubscribe(); } catch (_) {}
    });
    submissionListenerUnsubs = [];
    submissionSnapshotSources.clear();
    submissionsListenerMode = mode;

    const refreshMergedSubmissions = () => {
        const merged = new Map();
        submissionSnapshotSources.forEach((rows) => {
            rows.forEach((row) => merged.set(row.id, row));
        });
        allSubmissions = Array.from(merged.values());
        restoreAuditDuplicateCorrections(allSubmissions);
        if (auditReconciliationSourceRows.length && auditReconciliationResult) {
            auditReconciliationResult = computeAuditReconciliationResult(auditReconciliationSourceRows);
        }
        if (auditPaidReconciliationSourceRows.length) {
            recomputeAuditPaidReconciliationResult();
        }
        if (auditDuplicateScanResult) {
            auditDuplicateScanResult = buildAuditDuplicateScanResult();
        }
        scheduleCurrentTabRender();
    };

    const attachListener = (sourceKey, sourceQuery) => {
        const unsubscribe = onSnapshot(sourceQuery, (snapshot) => {
            submissionSnapshotSources.set(
                sourceKey,
                snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }))
            );
            refreshMergedSubmissions();
        }, () => {
            showNotification('Failed to load applications', 'error');
        });
        submissionListenerUnsubs.push(unsubscribe);
    };

    if (full) {
        attachListener('all', collection(db, 'submissions'));
        return;
    }

    attachListener(
        'payment-lifecycle',
        query(
            collection(db, 'submissions'),
            where('status', 'in', ['sent_to_pfa', 'rsa_submitted', 'paid', 'cleared'])
        )
    );
    attachListener(
        'audit-requests',
        query(
            collection(db, 'submissions'),
            where('auditCommissionStatus', 'in', ['pending', 'rejected'])
        )
    );
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
        if (role !== 'reports_monitoring' && role !== 'audit') {
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
        switchTab(getInitialAuditTab());
    } catch (_) {
        showNotification('Could not validate session', 'error');
        window.location.href = 'index.html';
    }
});
