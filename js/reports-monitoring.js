import { auth, db } from './firebase-config.js?v=20260625c';
import { performAppLogout } from './shared/logout.js?v=20260625b';
import {
    collection,
    query,
    where,
    addDoc,
    doc,
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
const AUDIT_DASHBOARD_TABS = ['overview', 'sent-to-pfa', 'paid', 'cleared', 'rejected', 'reconciliation', 'profile', 'help'];
const AUDIT_BULK_CLEAR_BATCH_SIZE = 200;

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
let auditPaidReconciliationSourceRows = [];
let auditPaidReconciliationResult = null;
let auditPaidReconciliationFileName = '';
let auditPaidReconciliationSelectedSubmissionIds = [];
let auditPaidReconciliationActiveResultsTab = 'matched';
let currentAuditRejectedScope = 'rejected';
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
const auditDuplicateScanBtn = document.getElementById('auditDuplicateScanBtn');
const auditReconciliationExportBtn = document.getElementById('auditReconciliationExportBtn');
const auditReconciliationClearBtn = document.getElementById('auditReconciliationClearBtn');
const auditReconciliationFileMeta = document.getElementById('auditReconciliationFileMeta');
const auditReconciliationSummary = document.getElementById('auditReconciliationSummary');
const auditDuplicateScanSummary = document.getElementById('auditDuplicateScanSummary');
const auditDuplicateScanWrap = document.getElementById('auditDuplicateScanWrap');
const auditDuplicateScanBody = document.getElementById('auditDuplicateScanBody');
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
    const ignoredStatuses = new Set(['draft', 'rejected', 'rejected_by_reviewer', 'rejected_by_rsa']);
    return allSubmissions.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return !ignoredStatuses.has(status);
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
    if (auditReconciliationFileInput) auditReconciliationFileInput.value = '';
    renderAuditReconciliation();
    renderAuditDuplicateScan();
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
    if (!auditDuplicateScanResult) {
        if (auditDuplicateScanSummary) {
            auditDuplicateScanSummary.style.display = 'none';
            auditDuplicateScanSummary.innerHTML = '';
        }
        if (auditDuplicateScanWrap) auditDuplicateScanWrap.style.display = 'none';
        if (auditDuplicateScanBody) auditDuplicateScanBody.innerHTML = '';
        return;
    }

    const groups = auditDuplicateScanResult.groups || [];
    if (auditDuplicateScanSummary) {
        auditDuplicateScanSummary.style.display = 'grid';
        auditDuplicateScanSummary.innerHTML = [
            { label: 'System Records Scanned', value: auditDuplicateScanResult.scannedCount },
            { label: 'Rejected Corrections Ignored', value: auditDuplicateScanResult.ignoredCorrectionCount || 0 },
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
        auditDuplicateScanBody.innerHTML = '<tr><td colspan="9" class="no-data">No duplicate applications found</td></tr>';
        return;
    }

    auditDuplicateScanBody.innerHTML = groups.map((group, groupIndex) => (
        group.rows.map((sub, rowIndex) => {
            const accountNumber = getCustomerAccountNumber(sub);
            const penNumber = getCustomerPenNumber(sub);
            const signalClass = group.strength === 'strong' ? 'audit-recon-status partial' : 'audit-recon-status info';
            const signalText = rowIndex === 0 ? group.key : 'Same duplicate group';
            const signalMeta = rowIndex === 0 ? `${group.rows.length} independent applications` : '';
            return `
                <tr>
                    <td>
                        <span class="${signalClass} audit-duplicate-signal">
                            <span>${escapeHtml(signalText)}</span>
                            ${signalMeta ? `<small>${escapeHtml(signalMeta)}</small>` : ''}
                        </span>
                    </td>
                    <td class="audit-duplicate-customer"><strong>${escapeHtml(sub.customerName || sub?.customerDetails?.name || 'Unknown')}</strong></td>
                    <td class="audit-duplicate-nowrap">${escapeHtml(accountNumber || '-')}</td>
                    <td class="audit-duplicate-nowrap">${escapeHtml(penNumber || '-')}</td>
                    <td><span class="audit-duplicate-status">${escapeHtml(statusLabel(sub.status || '-'))}</span></td>
                    <td>${escapeHtml(getUserDisplayName(sub.uploadedBy || ''))}</td>
                    <td class="audit-duplicate-nowrap">${escapeHtml(formatDate(getSubmissionOriginalUploadAt(sub)))}</td>
                    <td><code class="audit-duplicate-id">${escapeHtml(sub.id || '-')}</code></td>
                    <td>
                        <div class="audit-duplicate-actions">
                            <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
                            <button type="button" class="action-btn audit-duplicate-view-btn" onclick="window.openMonitoringApplicationTrack('${sub.id}')"><i class="fas fa-route"></i> Track</button>
                            <button type="button" class="action-btn audit-duplicate-delete-btn" onclick="window.deleteAuditDuplicateApplication('${sub.id}')"><i class="fas fa-trash"></i> Delete</button>
                        </div>
                    </td>
                </tr>
                ${rowIndex === group.rows.length - 1 && groupIndex < groups.length - 1 ? '<tr><td colspan="9" style="height:8px;background:#f8fafc;"></td></tr>' : ''}
            `;
        }).join('')
    )).join('');
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
    if (auditReconciliationClearBtn) auditReconciliationClearBtn.disabled = !selectedFile && !auditReconciliationResult;
    if (auditReconciliationExportBtn) auditReconciliationExportBtn.disabled = !auditReconciliationResult;

    if (!auditReconciliationResult) {
        if (auditReconciliationSummary) {
            auditReconciliationSummary.style.display = 'none';
            auditReconciliationSummary.innerHTML = '';
        }
        if (auditReconciliationMatchedWrap) auditReconciliationMatchedWrap.style.display = 'none';
        if (auditReconciliationMatchedBody) auditReconciliationMatchedBody.innerHTML = '';
        if (auditReconciliationUnmatchedWrap) auditReconciliationUnmatchedWrap.style.display = 'none';
        if (auditReconciliationUnmatchedBody) auditReconciliationUnmatchedBody.innerHTML = '';
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
    if (!auditReconciliationResult) throw new Error('Run reconciliation first.');
    if (!window.ExcelJS) throw new Error('Excel library is not available right now.');
    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Audit Dashboard';
    workbook.created = new Date();

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
                ? `<button class="action-btn" onclick="window.openMonitoringApplicationDetails('${sub.id}')"><i class="fas fa-eye"></i> View</button>
                   <button class="action-btn audit-clear-btn" onclick="window.clearAuditPayment('${sub.id}')"><i class="fas fa-circle-check"></i> Clear</button>`
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
        renderAuditPaidReconciliation();
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
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Audit';
}

function ensureDataForTab(tabId) {
    const dataTabs = ['overview', 'sent-to-pfa', 'paid', 'cleared', 'rejected', 'reconciliation'];
    if (dataTabs.includes(tabId)) loadUsers();
    if (dataTabs.includes(tabId)) loadSubmissions({ full: tabId === 'reconciliation' });
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

async function deleteAuditDuplicateApplication(submissionId) {
    const sub = allSubmissions.find((item) => item.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'warning');
        return;
    }

    const customerName = sub.customerName || sub?.customerDetails?.name || 'this application';
    const confirmed = window.confirm(`Delete duplicate application for "${customerName}"?\n\nThis will permanently remove this application record from the system.`);
    if (!confirmed) return;

    try {
        const batch = writeBatch(db);
        batch.delete(doc(db, 'submissions', submissionId));
        getSubmissionUniqueKeyRefs(sub).forEach((keyRef) => batch.delete(keyRef));
        await batch.commit();

        await addDoc(collection(db, 'audit'), {
            action: 'audit_duplicate_application_deleted',
            submissionId,
            customerName,
            accountNumber: getCustomerAccountNumber(sub),
            penNo: getCustomerPenNumber(sub),
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        }).catch(() => {});

        allSubmissions = allSubmissions.filter((item) => item.id !== submissionId);
        auditDuplicateScanResult = buildAuditDuplicateScanResult();
        renderAuditDuplicateScan();
        showNotification('Duplicate application deleted.', 'success');
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
    const confirmed = window.confirm(
        shouldFreeze
            ? `Freeze ${sub.customerName || 'this application'}? The uploader will not be able to resubmit, dissolve, chat, or make other changes until Audit unfreezes it.`
            : `Unfreeze ${sub.customerName || 'this application'}? The uploader will be able to act on it again.`
    );
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
    auditDuplicateScanBtn?.addEventListener('click', () => {
        const originalHtml = auditDuplicateScanBtn.innerHTML;
        auditDuplicateScanBtn.disabled = true;
        auditDuplicateScanBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Scanning...';
        try {
            auditDuplicateScanResult = buildAuditDuplicateScanResult();
            renderAuditDuplicateScan();
            showNotification(`Duplicate scan complete. ${auditDuplicateScanResult.groups.length} group(s) found.`, auditDuplicateScanResult.groups.length ? 'warning' : 'success');
        } catch (error) {
            auditDuplicateScanResult = null;
            renderAuditDuplicateScan();
            showNotification(error?.message || 'Failed to scan duplicate applications.', 'error');
        } finally {
            auditDuplicateScanBtn.disabled = false;
            auditDuplicateScanBtn.innerHTML = originalHtml;
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
            auditReconciliationExportBtn.disabled = !auditReconciliationResult;
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
