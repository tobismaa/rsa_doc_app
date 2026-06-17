import { auth, db } from './firebase-config.js';
import { EMAIL_API_BASE_URL } from './email-api-config.js';
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
import {
    getSubmissionReviewEntryAt,
    getSubmissionApprovalEntryAt,
    getSubmissionRsaEntryAt,
    getSubmissionFinalSubmissionEntryAt,
    getSubmissionPaymentEntryAt,
    getSubmissionPaidEntryAt,
    getSubmissionClearedEntryAt
} from './shared/submission-stage.js?v=20260610b';
import { clearSystemSettingsCache, getDefaultSystemSettings, getSystemSettings, normalizeAgentBankOptions } from './shared/system-settings.js?v=20260615a';
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
let currentPfaAddresses = {};
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
let settingsModalSourceDropdownTab = '';
let currentScheduledReportPreview = null;
let scheduledReportConfirmResolver = null;
let lastScheduledReportLogSnapshot = null;
let currentScheduledReportSendTab = 'daily';
let currentScheduledReportDownloadTab = 'daily';

const OUTSTANDING_REPORT_OPTIONS = [
    { id: 'all', label: 'All Dashboards' },
    { id: 'uploader', label: 'Uploader' },
    { id: 'reviewer', label: 'Reviewer' },
    { id: 'rsa', label: 'RSA' },
    { id: 'payment', label: 'Payment' }
];
const REPORT_INCEPTION_START_DATE = '1900-01-01';
const REPORT_INCEPTION_LABEL = 'From Inception';

function parseForceLogoutDurationInput(value) {
    const text = String(value || '').trim().toLowerCase();
    if (/^\d+\s*[smh]$/.test(text)) return text.replace(/\s+/g, '');
    if (/^\d+(\.\d+)?$/.test(text)) return `${text}m`;
    return '';
}

function forceLogoutDurationToSeconds(value) {
    const normalized = parseForceLogoutDurationInput(value);
    const match = normalized.match(/^(\d+(?:\.\d+)?)([smh])$/);
    if (!match) return 0;
    const amount = Number(match[1]);
    const unit = match[2];
    if (!Number.isFinite(amount) || amount <= 0) return 0;
    if (unit === 's') return amount;
    if (unit === 'm') return amount * 60;
    if (unit === 'h') return amount * 3600;
    return 0;
}

function describeForceLogoutDuration(value) {
    const normalized = parseForceLogoutDurationInput(value);
    const match = normalized.match(/^(\d+(?:\.\d+)?)([smh])$/);
    if (!match) return normalized || '11m';
    const amount = Number(match[1]);
    const unit = match[2];
    if (unit === 's') return `${amount} second${amount === 1 ? '' : 's'}`;
    if (unit === 'm') return `${amount} minute${amount === 1 ? '' : 's'}`;
    return `${amount} hour${amount === 1 ? '' : 's'}`;
}

function getEmailApiBaseUrl() {
    const runtime = String(window.__EMAIL_API_BASE_URL__ || '').trim();
    const configured = runtime || String(EMAIL_API_BASE_URL || '').trim();
    if (!configured || configured.includes('YOUR-RENDER-URL')) return '';
    return configured.replace(/\/+$/, '');
}

function humanizeScheduledReportTrigger(trigger = '') {
    const text = String(trigger || '').trim();
    if (!text) return 'Unknown';
    if (text === 'auto') return 'Automatic daily send';
    if (text === 'auto_weekend_rollup') return 'Automatic weekend rollup';
    if (text.startsWith('manual:')) return `Manual send by ${text.slice(7) || 'admin'}`;
    if (text.startsWith('manual_custom')) return 'Manual custom send';
    return text.replace(/_/g, ' ');
}

function buildScheduledReportMiniReport(status = {}) {
    if (!status || typeof status !== 'object') return '';
    const summary = status.todaySummary || {};
    const lastRun = status.lastRun || null;
    const todayRuns = Array.isArray(status.todayRuns) ? status.todayRuns : [];
    const statusTone = summary.hasSentForReportDate ? '#166534' : '#92400e';
    const statusText = summary.hasSentForReportDate
        ? `Yes, ${String(status.reportDateKey || '').trim() || 'today'} has been sent.`
        : `No successful send yet for ${String(status.reportDateKey || '').trim() || 'today'}.`;
    const summaryCards = [
        ['Report day sent', statusText],
        ['Today runs', String(Number(summary.totalRuns || 0))],
        ['Manual runs', String(Number(summary.manualRuns || 0))],
        ['Auto runs', String(Number(summary.autoRuns || 0))],
        ['Resends', String(Number(summary.resendRuns || 0))]
    ];
    const recentRunsMarkup = todayRuns.length
        ? todayRuns.slice(0, 6).map((run) => `
            <div style="padding:10px 12px;border:1px solid #dbe6f2;border-radius:10px;background:#fff;">
                <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;">
                    <strong style="color:#0f3b67;">${escapeHtml(run.reportLabel || run.reportDateKey || run.runKey || 'Scheduled report')}</strong>
                    <span style="font-size:12px;color:#64748b;">${escapeHtml(run.eventTime || '-')}</span>
                </div>
                <div style="margin-top:6px;font-size:12px;color:#334155;">${escapeHtml(humanizeScheduledReportTrigger(run.trigger))}</div>
                <div style="margin-top:4px;font-size:12px;color:#475569;">Status: ${escapeHtml(run.status || '-')} | Sent: ${Number(run.sentCount || 0)} | Failed: ${Number(run.failedCount || 0)}</div>
                ${run.attachmentFileName ? `<div style="margin-top:4px;font-size:12px;color:#475569;">Excel: ${escapeHtml(run.attachmentFileName)}</div>` : ''}
                ${run.error ? `<div style="margin-top:4px;font-size:12px;color:#991b1b;">Error: ${escapeHtml(run.error)}</div>` : ''}
            </div>
        `).join('')
        : '<div style="padding:10px 12px;border:1px dashed #cbd5e1;border-radius:10px;background:#fff;color:#64748b;">No send activity logged yet for today.</div>';

    return `
        <div style="margin-top:12px;padding-top:12px;border-top:1px solid rgba(148,163,184,0.3);">
            <div style="font-weight:700;color:#0f3b67;margin-bottom:8px;">Daily Send Report</div>
            <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin-bottom:10px;">
                ${summaryCards.map(([label, value], index) => `
                    <div style="padding:10px 12px;border:1px solid #dbe6f2;border-radius:10px;background:#fff;">
                        <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.04em;color:#64748b;">${escapeHtml(label)}</div>
                        <div style="margin-top:4px;font-size:${index === 0 ? '13px' : '18px'};font-weight:700;color:${index === 0 ? statusTone : '#0f172a'};">${escapeHtml(value)}</div>
                    </div>
                `).join('')}
            </div>
            ${lastRun ? `
                <div style="padding:10px 12px;border:1px solid #dbe6f2;border-radius:10px;background:#ffffff;margin-bottom:10px;">
                    <div style="font-weight:700;color:#0f3b67;margin-bottom:6px;">Latest report event</div>
                    <div style="font-size:12px;color:#334155;">Trigger: ${escapeHtml(humanizeScheduledReportTrigger(lastRun.trigger))}</div>
                    <div style="font-size:12px;color:#334155;">Configured send time: ${escapeHtml(lastRun.sendTime || status.sendTime || '-')}</div>
                    <div style="font-size:12px;color:#334155;">Excel file: ${escapeHtml(lastRun.attachmentFileName || 'Not recorded')}</div>
                    <div style="font-size:12px;color:#334155;">Completed: ${escapeHtml(lastRun.eventTime || '-')}</div>
                    ${lastRun.error ? `<div style="font-size:12px;color:#991b1b;">Error: ${escapeHtml(lastRun.error)}</div>` : ''}
                </div>
            ` : ''}
            <div style="font-weight:700;color:#0f3b67;margin-bottom:8px;">Today's activity</div>
            <div style="display:grid;gap:8px;">${recentRunsMarkup}</div>
        </div>
    `;
}

function setScheduledReportSendStatus(message = '', type = 'info', details = [], status = null) {
    const host = document.getElementById('scheduledReportSendStatus');
    if (!host) return;
    const safeType = ['success', 'error', 'warning', 'info'].includes(type) ? type : 'info';
    const toneMap = {
        success: { border: '#86efac', background: '#f0fdf4', color: '#166534', title: 'Sender Ready' },
        error: { border: '#fca5a5', background: '#fef2f2', color: '#991b1b', title: 'Sender Error' },
        warning: { border: '#fcd34d', background: '#fffbeb', color: '#92400e', title: 'Sender Attention' },
        info: { border: '#bfdbfe', background: '#eff6ff', color: '#1d4ed8', title: 'Sender Info' }
    };
    if (!String(message || '').trim() && (!Array.isArray(details) || !details.length) && !status) {
        host.style.display = 'none';
        host.innerHTML = '';
        return;
    }
    const tone = toneMap[safeType];
    host.style.display = 'block';
    host.style.borderColor = tone.border;
    host.style.background = tone.background;
    host.style.color = tone.color;
    host.innerHTML = `
        <div style="font-weight:700;margin-bottom:${details.length ? '6px' : '0'};">${escapeHtml(tone.title)}</div>
        <div>${escapeHtml(message)}</div>
        ${details.length ? `<div style="margin-top:8px;">${details.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</div>` : ''}
        ${status ? buildScheduledReportMiniReport(status) : ''}
    `;
}

function setScheduledReportSendModalStatus(message = '', type = 'info', details = []) {
    const host = document.getElementById('scheduledReportSendModalStatus');
    if (!host) return;
    const safeType = ['success', 'error', 'warning', 'info'].includes(type) ? type : 'info';
    const toneMap = {
        success: { border: '#86efac', background: '#f0fdf4', color: '#166534', title: 'Report Ready' },
        error: { border: '#fca5a5', background: '#fef2f2', color: '#991b1b', title: 'Report Error' },
        warning: { border: '#fcd34d', background: '#fffbeb', color: '#92400e', title: 'Report Attention' },
        info: { border: '#bfdbfe', background: '#eff6ff', color: '#1d4ed8', title: 'Report Info' }
    };
    if (!String(message || '').trim() && (!Array.isArray(details) || !details.length)) {
        host.style.display = 'none';
        host.innerHTML = '';
        return;
    }
    const tone = toneMap[safeType];
    host.style.display = 'block';
    host.style.borderColor = tone.border;
    host.style.background = tone.background;
    host.style.color = tone.color;
    host.innerHTML = `
        <div style="font-weight:700;margin-bottom:${details.length ? '6px' : '0'};">${escapeHtml(tone.title)}</div>
        <div>${escapeHtml(message)}</div>
        ${details.length ? `<div style="margin-top:8px;">${details.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</div>` : ''}
    `;
}

function setOutstandingReportDownloadStatus(message = '', type = 'info', details = []) {
    const host = document.getElementById('scheduledOutstandingDownloadStatus');
    if (!host) return;
    const safeType = ['success', 'error', 'warning', 'info'].includes(type) ? type : 'info';
    const toneMap = {
        success: { border: '#86efac', background: '#f0fdf4', color: '#166534', title: 'Outstanding Report Ready' },
        error: { border: '#fca5a5', background: '#fef2f2', color: '#991b1b', title: 'Outstanding Report Error' },
        warning: { border: '#fcd34d', background: '#fffbeb', color: '#92400e', title: 'Outstanding Report Attention' },
        info: { border: '#bfdbfe', background: '#eff6ff', color: '#1d4ed8', title: 'Outstanding Report Info' }
    };
    if (!String(message || '').trim() && (!Array.isArray(details) || !details.length)) {
        host.style.display = 'none';
        host.innerHTML = '';
        return;
    }
    const tone = toneMap[safeType];
    host.style.display = 'block';
    host.style.borderColor = tone.border;
    host.style.background = tone.background;
    host.style.color = tone.color;
    host.innerHTML = `
        <div style="font-weight:700;margin-bottom:${details.length ? '6px' : '0'};">${escapeHtml(tone.title)}</div>
        <div>${escapeHtml(message)}</div>
        ${details.length ? `<div style="margin-top:8px;">${details.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</div>` : ''}
    `;
}

function openScheduledReportSendModal() {
    const modal = document.getElementById('scheduledReportSendModal');
    if (!modal) return;
    modal.classList.add('active');
    const previousDay = getPreviousScheduledReportDateKey();
    const manualDate = document.getElementById('scheduledReportManualDate');
    const rangeStart = document.getElementById('scheduledReportRangeStartDate');
    const rangeEnd = document.getElementById('scheduledReportRangeEndDate');
    const inceptionEnd = document.getElementById('scheduledReportInceptionEndDate');
    const mode = document.getElementById('scheduledReportSendMode');
    const outstandingDate = document.getElementById('scheduledOutstandingReportDate');
    const outstandingRangeStart = document.getElementById('scheduledOutstandingRangeStartDate');
    const outstandingRangeEnd = document.getElementById('scheduledOutstandingRangeEndDate');
    const outstandingInceptionEnd = document.getElementById('scheduledOutstandingInceptionEndDate');
    const outstandingMode = document.getElementById('scheduledOutstandingSendMode');
    if (manualDate && !manualDate.value) manualDate.value = previousDay;
    if (rangeStart && !rangeStart.value) rangeStart.value = previousDay;
    if (rangeEnd && !rangeEnd.value) rangeEnd.value = previousDay;
    if (inceptionEnd && !inceptionEnd.value) inceptionEnd.value = previousDay;
    if (outstandingDate && !outstandingDate.value) outstandingDate.value = previousDay;
    if (outstandingRangeStart && !outstandingRangeStart.value) outstandingRangeStart.value = previousDay;
    if (outstandingRangeEnd && !outstandingRangeEnd.value) outstandingRangeEnd.value = previousDay;
    if (outstandingInceptionEnd && !outstandingInceptionEnd.value) outstandingInceptionEnd.value = previousDay;
    if (mode && !mode.value) mode.value = 'single_day';
    if (outstandingMode && !outstandingMode.value) outstandingMode.value = 'single_day';
    currentScheduledReportSendTab = 'daily';
    renderScheduledReportSendTabState();
    updateScheduledReportSendModeVisibility();
    updateOutstandingReportSendModeVisibility();
    setScheduledReportSendModalStatus();
}

function closeScheduledReportSendModal() {
    document.getElementById('scheduledReportSendModal')?.classList.remove('active');
    setScheduledReportSendModalStatus();
}

function openScheduledReportLogsModal() {
    document.getElementById('scheduledReportLogsModal')?.classList.add('active');
    renderScheduledReportLogs(Boolean(lastScheduledReportLogSnapshot)).catch(() => {});
}

function closeScheduledReportLogsModal() {
    document.getElementById('scheduledReportLogsModal')?.classList.remove('active');
}

function openScheduledReportConfirmModal({
    title = 'Confirm Action',
    heroTitle = 'Please confirm',
    message = 'Do you want to continue?',
    note = 'This action will continue with the selected report operation.',
    confirmLabel = 'Continue'
} = {}) {
    const modal = document.getElementById('scheduledReportConfirmModal');
    if (!modal) return Promise.resolve(false);
    document.getElementById('scheduledReportConfirmModalTitle').innerHTML = `<i class="fas fa-circle-question"></i> ${escapeHtml(title)}`;
    const heroTitleEl = document.getElementById('scheduledReportConfirmModalHeroTitle');
    const messageEl = document.getElementById('scheduledReportConfirmModalMessage');
    const noteEl = document.getElementById('scheduledReportConfirmModalNote');
    const confirmBtn = document.getElementById('confirmScheduledReportConfirmModalBtn');
    if (heroTitleEl) heroTitleEl.textContent = heroTitle;
    if (messageEl) messageEl.textContent = message;
    if (noteEl) noteEl.textContent = note;
    if (confirmBtn) confirmBtn.innerHTML = `<i class="fas fa-paper-plane"></i> ${escapeHtml(confirmLabel)}`;
    modal.classList.add('active');
    return new Promise((resolve) => {
        scheduledReportConfirmResolver = resolve;
    });
}

function closeScheduledReportConfirmModal(confirmed = false) {
    document.getElementById('scheduledReportConfirmModal')?.classList.remove('active');
    const resolver = scheduledReportConfirmResolver;
    scheduledReportConfirmResolver = null;
    if (typeof resolver === 'function') resolver(confirmed === true);
}

function updateScheduledReportSendModeVisibility() {
    const mode = String(document.getElementById('scheduledReportSendMode')?.value || 'single_day').trim();
    const singleWrap = document.getElementById('scheduledReportSingleDateWrap');
    const rangeWrap = document.getElementById('scheduledReportRangeWrap');
    const inceptionWrap = document.getElementById('scheduledReportInceptionWrap');
    if (singleWrap) singleWrap.style.display = mode === 'single_day' ? '' : 'none';
    if (rangeWrap) rangeWrap.style.display = mode === 'date_range' ? 'grid' : 'none';
    if (inceptionWrap) inceptionWrap.style.display = mode === 'from_inception' ? 'grid' : 'none';
}

function updateOutstandingReportSendModeVisibility() {
    const mode = String(document.getElementById('scheduledOutstandingSendMode')?.value || 'single_day').trim();
    const singleWrap = document.getElementById('scheduledOutstandingSingleDateWrap');
    const rangeWrap = document.getElementById('scheduledOutstandingRangeWrap');
    const inceptionWrap = document.getElementById('scheduledOutstandingInceptionWrap');
    if (singleWrap) singleWrap.style.display = mode === 'single_day' ? '' : 'none';
    if (rangeWrap) rangeWrap.style.display = mode === 'date_range' ? 'grid' : 'none';
    if (inceptionWrap) inceptionWrap.style.display = mode === 'from_inception' ? 'grid' : 'none';
}

function updateOutstandingReportDownloadModeVisibility() {
    const mode = String(document.getElementById('scheduledOutstandingDownloadMode')?.value || 'single_day').trim();
    const singleWrap = document.getElementById('scheduledOutstandingDownloadSingleDateWrap');
    const rangeWrap = document.getElementById('scheduledOutstandingDownloadRangeWrap');
    const inceptionWrap = document.getElementById('scheduledOutstandingDownloadInceptionWrap');
    if (singleWrap) singleWrap.style.display = mode === 'single_day' ? '' : 'none';
    if (rangeWrap) rangeWrap.style.display = mode === 'date_range' ? 'grid' : 'none';
    if (inceptionWrap) inceptionWrap.style.display = mode === 'from_inception' ? 'grid' : 'none';
}

function renderScheduledReportSendTabState() {
    const dailyBtn = document.getElementById('scheduledReportSendDailyTabBtn');
    const outstandingBtn = document.getElementById('scheduledReportSendOutstandingTabBtn');
    const dailySection = document.getElementById('scheduledReportSendDailySection');
    const outstandingSection = document.getElementById('scheduledReportSendOutstandingSection');
    const isOutstanding = currentScheduledReportSendTab === 'outstanding';
    if (dailyBtn) {
        dailyBtn.style.background = isOutstanding ? '' : '#003366';
        dailyBtn.style.color = isOutstanding ? '' : '#fff';
        dailyBtn.style.border = isOutstanding ? '' : 'none';
    }
    if (outstandingBtn) {
        outstandingBtn.style.background = isOutstanding ? '#003366' : '';
        outstandingBtn.style.color = isOutstanding ? '#fff' : '';
        outstandingBtn.style.border = isOutstanding ? 'none' : '';
    }
    if (dailySection) dailySection.style.display = isOutstanding ? 'none' : '';
    if (outstandingSection) outstandingSection.style.display = isOutstanding ? '' : 'none';
}

function renderScheduledReportDownloadTabState() {
    const dailyBtn = document.getElementById('scheduledReportDownloadDailyTabBtn');
    const outstandingBtn = document.getElementById('scheduledReportDownloadOutstandingTabBtn');
    const dailySection = document.getElementById('scheduledReportDownloadDailySection');
    const outstandingSection = document.getElementById('scheduledReportDownloadOutstandingSection');
    const isOutstanding = currentScheduledReportDownloadTab === 'outstanding';
    if (dailyBtn) {
        dailyBtn.style.background = isOutstanding ? '' : '#003366';
        dailyBtn.style.color = isOutstanding ? '' : '#fff';
        dailyBtn.style.border = isOutstanding ? '' : 'none';
    }
    if (outstandingBtn) {
        outstandingBtn.style.background = isOutstanding ? '#003366' : '';
        outstandingBtn.style.color = isOutstanding ? '#fff' : '';
        outstandingBtn.style.border = isOutstanding ? 'none' : '';
    }
    if (dailySection) dailySection.style.display = isOutstanding ? 'none' : '';
    if (outstandingSection) outstandingSection.style.display = isOutstanding ? '' : 'none';
}

function buildManualScheduledReportPayload() {
    if (currentScheduledReportSendTab === 'outstanding') {
        return buildOutstandingReportSelection({
            dashboardElementId: 'scheduledOutstandingDashboardSelect',
            modeElementId: 'scheduledOutstandingSendMode',
            singleDateElementId: 'scheduledOutstandingReportDate',
            rangeStartElementId: 'scheduledOutstandingRangeStartDate',
            rangeEndElementId: 'scheduledOutstandingRangeEndDate',
            inceptionEndElementId: 'scheduledOutstandingInceptionEndDate'
        });
    }
    const mode = String(document.getElementById('scheduledReportSendMode')?.value || 'single_day').trim();
    if (mode === 'date_range') {
        const rangeStartDateKey = String(document.getElementById('scheduledReportRangeStartDate')?.value || '').trim();
        const rangeEndDateKey = String(document.getElementById('scheduledReportRangeEndDate')?.value || '').trim();
        if (!rangeStartDateKey || !rangeEndDateKey) {
            throw new Error('Choose both start date and end date for the report range.');
        }
        return {
            mode: 'date_range',
            payload: { rangeStartDateKey, rangeEndDateKey },
            label: `${rangeStartDateKey} to ${rangeEndDateKey}`
        };
    }
    if (mode === 'from_inception') {
        const rangeEndDateKey = String(document.getElementById('scheduledReportInceptionEndDate')?.value || '').trim();
        if (!rangeEndDateKey) {
            throw new Error('Choose the end date for the inception report.');
        }
        return {
            mode: 'from_inception',
            payload: { rangeStartDateKey: REPORT_INCEPTION_START_DATE, rangeEndDateKey },
            label: `${REPORT_INCEPTION_LABEL} to ${rangeEndDateKey}`
        };
    }
    const reportDateKey = String(document.getElementById('scheduledReportManualDate')?.value || '').trim();
    if (!reportDateKey) {
        throw new Error('Choose the report date you want to send.');
    }
    return {
        mode: 'single_day',
        payload: { reportDateKey },
        label: reportDateKey
    };
}

function buildOutstandingReportSelection({
    dashboardElementId,
    modeElementId,
    singleDateElementId,
    rangeStartElementId,
    rangeEndElementId,
    inceptionEndElementId
} = {}) {
    const dashboard = String(document.getElementById(dashboardElementId)?.value || 'all').trim().toLowerCase() || 'all';
    const option = OUTSTANDING_REPORT_OPTIONS.find((item) => item.id === dashboard) || OUTSTANDING_REPORT_OPTIONS[0];
    const mode = String(document.getElementById(modeElementId)?.value || 'single_day').trim();
    if (mode === 'date_range') {
        const rangeStartDateKey = String(document.getElementById(rangeStartElementId)?.value || '').trim() || getPreviousScheduledReportDateKey();
        const rangeEndDateKey = String(document.getElementById(rangeEndElementId)?.value || '').trim() || getPreviousScheduledReportDateKey();
        return {
            mode: 'outstanding_range',
            payload: { rangeStartDateKey, rangeEndDateKey, outstandingDashboard: dashboard },
            label: `${option.label} Outstanding - ${rangeStartDateKey} to ${rangeEndDateKey}`,
            reportOptions: {
                reportDateKey: '',
                rangeStartDateKey,
                rangeEndDateKey,
                dashboard
            }
        };
    }
    if (mode === 'from_inception') {
        const rangeEndDateKey = String(document.getElementById(inceptionEndElementId)?.value || '').trim() || getPreviousScheduledReportDateKey();
        return {
            mode: 'outstanding_inception',
            payload: { rangeStartDateKey: REPORT_INCEPTION_START_DATE, rangeEndDateKey, outstandingDashboard: dashboard },
            label: `${option.label} Outstanding - ${REPORT_INCEPTION_LABEL} to ${rangeEndDateKey}`,
            reportOptions: {
                reportDateKey: '',
                rangeStartDateKey: REPORT_INCEPTION_START_DATE,
                rangeEndDateKey,
                dashboard
            }
        };
    }
    const reportDateKey = String(document.getElementById(singleDateElementId)?.value || '').trim() || getPreviousScheduledReportDateKey();
    return {
        mode: 'outstanding',
        payload: { reportDateKey, outstandingDashboard: dashboard },
        label: `${option.label} Outstanding - ${reportDateKey}`,
        reportOptions: {
            reportDateKey,
            rangeStartDateKey: '',
            rangeEndDateKey: '',
            dashboard
        }
    };
}

function buildOutstandingDownloadSelection() {
    return buildOutstandingReportSelection({
        dashboardElementId: 'scheduledOutstandingDownloadDashboardSelect',
        modeElementId: 'scheduledOutstandingDownloadMode',
        singleDateElementId: 'scheduledOutstandingDownloadDate',
        rangeStartElementId: 'scheduledOutstandingDownloadRangeStartDate',
        rangeEndElementId: 'scheduledOutstandingDownloadRangeEndDate',
        inceptionEndElementId: 'scheduledOutstandingDownloadInceptionEndDate'
    });
}

async function sendSelectedScheduledReport(selection, {
    setStatus = () => {},
    successMessage = 'Selected report sent successfully.',
    busyMessage = null
} = {}) {
    const status = await getScheduledReportBackendStatus();
    if (status.enabled !== true) {
        setStatus('Daily report email is disabled in settings.', 'warning', describeScheduledReportStatus(status));
        return { ok: false, skipped: true, reason: 'scheduled-report-disabled' };
    }
    if (!Array.isArray(status.recipients) || !status.recipients.length) {
        setStatus('No recipients are configured for scheduled report emails.', 'warning', describeScheduledReportStatus(status));
        return { ok: false, skipped: true, reason: 'no-recipients' };
    }

    const apiBaseUrl = getEmailApiBaseUrl();
    const idToken = await currentUser.getIdToken(true);
    setStatus(busyMessage || `Sending report for ${selection.label}...`, 'info', describeScheduledReportStatus(status));

    let requestResult = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/send-now`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${idToken}`
        },
        body: JSON.stringify(selection.payload)
    }, 45000);

    if (requestResult.response.status === 409 || String(requestResult.data?.error || '').trim() === 'already-sent') {
        const confirmed = await openScheduledReportConfirmModal({
            title: 'Report Already Sent',
            heroTitle: 'Custom report was already sent',
            message: `A report for ${selection.label} has already been sent. Do you still want to resend it?`,
            note: 'Continuing will resend this same report to the configured recipient list.',
            confirmLabel: 'Resend Report'
        });
        if (!confirmed) {
            setStatus(`Manual resend cancelled for ${selection.label}.`, 'info', describeScheduledReportStatus(status));
            return { ok: false, skipped: true, reason: 'cancelled' };
        }
        requestResult = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/send-now`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${idToken}`
            },
            body: JSON.stringify({ ...selection.payload, forceResend: true })
        }, 45000);
    }

    const result = requestResult.data || {};
    if (!requestResult.response.ok || result?.ok === false) {
        throw new Error(String(result?.error || 'Failed to send selected report.'));
    }

    const details = [
        `Report sent: ${String(result?.reportLabel || selection.label).trim() || selection.label}`,
        `Excel file: ${String(result?.attachmentFileName || '-').trim() || '-'}`,
        `Sent count: ${Number(result?.sentCount || 0)}`,
        `Failed count: ${Number(result?.failedCount || 0)}`
    ];
    const refreshedStatus = await getScheduledReportBackendStatus().catch(() => null);
    setStatus(successMessage, 'success', details);
    setScheduledReportSendStatus(successMessage, 'success', details, refreshedStatus);
    return { ok: true, result, details, status: refreshedStatus };
}

async function sendPrebuiltScheduledWorkbook({
    report,
    reportLabel,
    setStatus = () => {},
    successMessage = 'Selected report sent successfully.',
    busyMessage = null,
    forceResend = false
} = {}) {
    const artifact = await buildDailyReportWorkbookArtifact(report);
    if (!artifact) return { ok: false, skipped: true, reason: 'no-artifact' };

    const status = await getScheduledReportBackendStatus();
    if (status.enabled !== true) {
        setStatus('Daily report email is disabled in settings.', 'warning', describeScheduledReportStatus(status));
        return { ok: false, skipped: true, reason: 'scheduled-report-disabled' };
    }
    if (!Array.isArray(status.recipients) || !status.recipients.length) {
        setStatus('No recipients are configured for scheduled report emails.', 'warning', describeScheduledReportStatus(status));
        return { ok: false, skipped: true, reason: 'no-recipients' };
    }

    const apiBaseUrl = getEmailApiBaseUrl();
    const idToken = await currentUser.getIdToken(true);
    const normalizedLabel = String(reportLabel || artifact.report?.reportLabel || '').trim() || 'Outstanding Report';
    setStatus(busyMessage || `Sending report for ${normalizedLabel}...`, 'info', describeScheduledReportStatus(status));

    const bodyPayload = {
        reportLabel: normalizedLabel,
        attachmentFileName: artifact.fileName,
        attachmentDataUri: `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${arrayBufferToBase64(artifact.buffer)}`,
        reportType: artifact.isOutstanding ? 'outstanding' : 'daily',
        outstandingDashboard: artifact.report?.selectedDashboard || '',
        reportDateKey: String(artifact.report?.dateKey || '').trim(),
        rangeStartDateKey: String(artifact.report?.rangeStartDateKey || '').trim(),
        rangeEndDateKey: String(artifact.report?.rangeEndDateKey || '').trim(),
        forceResend: forceResend === true
    };

    let requestResult = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/send-workbook`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${idToken}`
        },
        body: JSON.stringify(bodyPayload)
    }, 60000);

    if (requestResult.response.status === 409 || String(requestResult.data?.error || '').trim() === 'already-sent') {
        const confirmed = await openScheduledReportConfirmModal({
            title: 'Report Already Sent',
            heroTitle: 'Custom report was already sent',
            message: `A report for ${normalizedLabel} has already been sent. Do you still want to resend it?`,
            note: 'Continuing will resend this exact workbook to the configured recipient list.',
            confirmLabel: 'Resend Report'
        });
        if (!confirmed) {
            setStatus(`Manual resend cancelled for ${normalizedLabel}.`, 'info', describeScheduledReportStatus(status));
            return { ok: false, skipped: true, reason: 'cancelled' };
        }
        bodyPayload.forceResend = true;
        requestResult = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/send-workbook`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${idToken}`
            },
            body: JSON.stringify(bodyPayload)
        }, 60000);
    }

    const result = requestResult.data || {};
    if (!requestResult.response.ok || result?.ok === false) {
        throw new Error(String(result?.error || 'Failed to send selected report.'));
    }

    const details = [
        `Report sent: ${String(result?.reportLabel || normalizedLabel).trim() || normalizedLabel}`,
        `Excel file: ${String(result?.attachmentFileName || artifact.fileName).trim() || artifact.fileName}`,
        `Sent count: ${Number(result?.sentCount || 0)}`,
        `Failed count: ${Number(result?.failedCount || 0)}`
    ];
    const refreshedStatus = await getScheduledReportBackendStatus().catch(() => null);
    setStatus(successMessage, 'success', details);
    setScheduledReportSendStatus(successMessage, 'success', details, refreshedStatus);
    return { ok: true, result, details, status: refreshedStatus };
}

function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    let binary = '';
    for (let index = 0; index < bytes.length; index += chunkSize) {
        const chunk = bytes.subarray(index, index + chunkSize);
        binary += String.fromCharCode.apply(null, Array.from(chunk));
    }
    return btoa(binary);
}

async function fetchJsonWithTimeout(url, options = {}, timeoutMs = 20000) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
        const response = await fetch(url, { ...options, signal: controller.signal });
        const data = await response.json().catch(() => ({}));
        return { response, data };
    } finally {
        clearTimeout(timer);
    }
}

async function getScheduledReportBackendStatus() {
    const apiBaseUrl = getEmailApiBaseUrl();
    if (!apiBaseUrl) {
        throw new Error('Scheduled report sender backend URL is not configured.');
    }
    if (!currentUser) {
        throw new Error('You must be signed in to verify sender status.');
    }
    const idToken = await currentUser.getIdToken(true);
    const { response, data } = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/status`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${idToken}`
        }
    }, 15000);
    if (!response.ok || data?.ok === false) {
        throw new Error(String(data?.error || 'Failed to check scheduled report sender status.'));
    }
    return data;
}

function describeScheduledReportStatus(status = {}) {
    const details = [
        `Report date: ${String(status.reportDateKey || '-').trim() || '-'}`,
        `Lagos time now: ${String(status.currentLagosTime || '-').trim() || '-'}`,
        `Configured send time: ${String(status.sendTime || '-').trim() || '-'}`,
        `Recipients loaded: ${Array.isArray(status.recipients) ? status.recipients.length : 0}`,
        `EmailJS backend config: ${status.emailJsConfigured === true ? 'Ready' : 'Missing or incomplete'}`
    ];
    const lastRun = status.lastRun || null;
    if (lastRun) {
        details.push(`Last run status: ${String(lastRun.status || '-').trim() || '-'}`);
        if (lastRun.trigger) details.push(`Last run trigger: ${String(lastRun.trigger || '-').trim() || '-'}`);
        if (lastRun.sendTime) details.push(`Last run configured time: ${String(lastRun.sendTime || '-').trim() || '-'}`);
        if (lastRun.sentCount !== undefined) details.push(`Last run sent count: ${Number(lastRun.sentCount || 0)}`);
        if (lastRun.failedCount !== undefined) details.push(`Last run failed count: ${Number(lastRun.failedCount || 0)}`);
    }
    return details;
}

function getScheduledReportLogTone(type = 'info') {
    const tones = {
        success: { border: '#86efac', background: '#f0fdf4', color: '#166534', title: 'Report Logs Ready' },
        error: { border: '#fca5a5', background: '#fef2f2', color: '#991b1b', title: 'Report Logs Error' },
        warning: { border: '#fcd34d', background: '#fffbeb', color: '#92400e', title: 'Report Logs Attention' },
        info: { border: '#bfdbfe', background: '#eff6ff', color: '#1d4ed8', title: 'Report Logs Info' }
    };
    return tones[type] || tones.info;
}

function setScheduledReportLogsBanner(message = '', type = 'info', details = []) {
    const host = document.getElementById('scheduledReportLogsBanner');
    if (!host) return;
    const text = String(message || '').trim();
    if (!text && (!Array.isArray(details) || !details.length)) {
        host.style.display = 'none';
        host.innerHTML = '';
        return;
    }
    const tone = getScheduledReportLogTone(type);
    host.style.display = 'block';
    host.style.borderColor = tone.border;
    host.style.background = tone.background;
    host.style.color = tone.color;
    host.innerHTML = `
        <div style="font-weight:700;margin-bottom:${details.length ? '6px' : '0'};">${escapeHtml(tone.title)}</div>
        <div>${escapeHtml(text)}</div>
        ${details.length ? `<div style="margin-top:8px;">${details.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</div>` : ''}
    `;
}

function setScheduledReportLogsInlineStatus(message = '', type = 'info') {
    const host = document.getElementById('scheduledReportLogsInlineStatus');
    if (!host) return;
    const text = String(message || '').trim();
    if (!text && !lastScheduledReportLogSnapshot) {
        host.style.display = 'none';
        host.innerHTML = '';
        return;
    }
    const tone = getScheduledReportLogTone(type);
    const snapshotMarkup = lastScheduledReportLogSnapshot ? buildScheduledReportMiniReport(lastScheduledReportLogSnapshot) : '';
    host.style.display = 'block';
    host.innerHTML = `
        <div style="padding:12px 14px;border-radius:12px;border:1px solid ${tone.border};background:${tone.background};color:${tone.color};font-size:13px;line-height:1.6;">
            ${text ? `<div style="font-weight:700;margin-bottom:6px;">${escapeHtml(text)}</div>` : ''}
            ${snapshotMarkup || '<div style="color:#64748b;">No scheduled report log snapshot available yet.</div>'}
        </div>
    `;
}

function buildScheduledReportLogSummaryCards(status = {}) {
    const summary = status.todaySummary || {};
    const lastRun = status.lastRun || null;
    const cards = [
        {
            tone: summary.hasSentForReportDate ? 'approved' : 'pending',
            label: 'Report Day Status',
            value: summary.hasSentForReportDate ? 'Sent' : 'Not Sent',
            icon: summary.hasSentForReportDate ? 'fa-circle-check' : 'fa-clock'
        },
        {
            tone: 'pending',
            label: 'Today Runs',
            value: String(Number(summary.totalRuns || 0)),
            icon: 'fa-list-check'
        },
        {
            tone: 'approved',
            label: 'Manual Runs',
            value: String(Number(summary.manualRuns || 0)),
            icon: 'fa-hand-pointer'
        },
        {
            tone: 'approved',
            label: 'Auto Runs',
            value: String(Number(summary.autoRuns || 0)),
            icon: 'fa-robot'
        },
        {
            tone: Number(summary.failedRuns || 0) > 0 ? 'rejected' : 'approved',
            label: 'Failed Runs',
            value: String(Number(summary.failedRuns || 0)),
            icon: Number(summary.failedRuns || 0) > 0 ? 'fa-triangle-exclamation' : 'fa-circle-check'
        },
        {
            tone: 'pending',
            label: 'Last Event',
            value: lastRun?.eventTime || '-',
            icon: 'fa-calendar-check'
        }
    ];
    return cards.map((card) => `
        <div class="stat-card ${card.tone}">
            <div class="stat-content">
                <div>
                    <div class="stat-label">${escapeHtml(card.label)}</div>
                    <div class="stat-value" style="font-size:${card.label === 'Last Event' ? '18px' : '28px'};">${escapeHtml(card.value)}</div>
                </div>
                <i class="fas ${card.icon} stat-icon"></i>
            </div>
        </div>
    `).join('');
}

function buildScheduledReportLogRows(runs = []) {
    if (!Array.isArray(runs) || !runs.length) {
        return '<tr><td colspan="7" class="no-data">No scheduled report activity has been logged yet.</td></tr>';
    }
    return runs.map((run) => {
        const statusText = String(run.status || '-').trim() || '-';
        const statusLower = statusText.toLowerCase();
        const badgeClass = statusLower === 'sent' || statusLower === 'partial'
            ? 'approved'
            : (statusLower === 'failed' ? 'rejected' : 'pending');
        const reportWindow = run.rangeStartDateKey && run.rangeEndDateKey && run.rangeStartDateKey !== run.rangeEndDateKey
            ? `${run.rangeStartDateKey} to ${run.rangeEndDateKey}`
            : (run.reportDateKey || '-');
        const recipientsText = `${Number(run.sentCount || 0)}/${Number(run.recipientsCount || 0)}`;
        return `
            <tr>
                <td>${escapeHtml(run.eventTime || '-')}</td>
                <td>
                    <strong>${escapeHtml(run.reportLabel || reportWindow)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">Key: ${escapeHtml(run.runKey || run.reportDateKey || '-')}</div>
                </td>
                <td>
                    <div>${escapeHtml(humanizeScheduledReportTrigger(run.trigger))}</div>
                    ${run.resendRequested ? '<div style="font-size:12px;color:#b45309;margin-top:4px;">Resend requested</div>' : ''}
                </td>
                <td><span class="status-badge ${badgeClass}">${escapeHtml(statusText)}</span></td>
                <td>
                    <div>Sent: ${Number(run.sentCount || 0)}</div>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">Failed: ${Number(run.failedCount || 0)} | Total: ${escapeHtml(recipientsText)}</div>
                    ${run.error ? `<div style="font-size:12px;color:#991b1b;margin-top:4px;">${escapeHtml(run.error)}</div>` : ''}
                </td>
                <td>${escapeHtml(run.attachmentFileName || '-')}</td>
                <td>${escapeHtml(run.sendTime || '-')}</td>
            </tr>
        `;
    }).join('');
}

function renderScheduledReportLogsFromStatus(status = {}) {
    const summaryHost = document.getElementById('scheduledReportLogsSummary');
    const body = document.getElementById('scheduledReportLogsTableBody');
    const updatedAt = document.getElementById('scheduledReportLogsUpdatedAt');
    if (summaryHost) summaryHost.innerHTML = buildScheduledReportLogSummaryCards(status);
    const runs = Array.isArray(status.recentRuns) && status.recentRuns.length
        ? status.recentRuns
        : (Array.isArray(status.todayRuns) ? status.todayRuns : []);
    if (body) body.innerHTML = buildScheduledReportLogRows(runs);
    if (updatedAt) updatedAt.textContent = `Last checked: ${status.currentLagosTime || '-'}`;
}

async function renderScheduledReportLogs(forceRefresh = false) {
    if (!forceRefresh && lastScheduledReportLogSnapshot) {
        renderScheduledReportLogsFromStatus(lastScheduledReportLogSnapshot);
        return;
    }
    const summaryHost = document.getElementById('scheduledReportLogsSummary');
    const body = document.getElementById('scheduledReportLogsTableBody');
    const updatedAt = document.getElementById('scheduledReportLogsUpdatedAt');
    if (summaryHost) summaryHost.innerHTML = '';
    if (body) body.innerHTML = '<tr><td colspan="7" class="no-data">Loading scheduled report logs...</td></tr>';
    if (updatedAt) updatedAt.textContent = 'Checking sender activity...';
    setScheduledReportLogsBanner('Checking scheduled report log history...', 'info');
    try {
        const status = await getScheduledReportBackendStatus();
        lastScheduledReportLogSnapshot = status;
        renderScheduledReportLogsFromStatus(status);
        setScheduledReportLogsInlineStatus('Report log snapshot updated.', 'success');
        const detailLines = [
            `Report day in focus: ${String(status.reportDateKey || '-').trim() || '-'}`,
            `Configured send time: ${String(status.sendTime || '-').trim() || '-'}`,
            `Recipients configured: ${Array.isArray(status.recipients) ? status.recipients.length : 0}`
        ];
        if (status.enabled !== true) {
            setScheduledReportLogsInlineStatus('Scheduled daily report email is disabled right now.', 'warning');
            setScheduledReportLogsBanner('Scheduled daily report email is currently disabled in settings.', 'warning', detailLines);
            return;
        }
        setScheduledReportLogsBanner('Scheduled report logs loaded successfully.', 'success', detailLines);
    } catch (error) {
        const message = String(error?.name === 'AbortError' ? 'Timed out while loading scheduled report logs.' : (error?.message || 'Failed to load scheduled report logs.'));
        if (body) body.innerHTML = `<tr><td colspan="7" class="no-data">${escapeHtml(message)}</td></tr>`;
        if (updatedAt) updatedAt.textContent = 'Last checked: failed';
        setScheduledReportLogsInlineStatus(message, 'error');
        setScheduledReportLogsBanner(message, 'error');
    }
}

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
    'Scheduled Report Emails': 'fa-paper-plane',
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
        return ['super_admin', 'admin', 'reports_monitoring', 'reviewer', 'rsa', 'payment'].includes(role) && status !== 'active';
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

function normalizeEmailList(values = []) {
    const seen = new Set();
    const emails = [];
    values.forEach((value) => {
        const email = normalizeEmail(value);
        if (!email || seen.has(email)) return;
        seen.add(email);
        emails.push(email);
    });
    return emails;
}

function getScheduledReportRecipients() {
    return normalizeEmailList(parseLinesTextarea('settingScheduledReportRecipients'));
}

function syncScheduledReportRecipientsTextarea(emails = []) {
    const field = document.getElementById('settingScheduledReportRecipients');
    if (!field) return;
    field.value = normalizeEmailList(emails).join('\n');
}

function getScheduledReportSelectableUsers(searchTerm = '') {
    const term = String(searchTerm || '').trim().toLowerCase();
    return allUsers
        .filter((user) => normalizeEmail(user?.email))
        .filter((user) => {
            if (!term) return true;
            const haystack = [
                user?.fullName,
                user?.email,
                user?.role,
                user?.department,
                user?.location
            ].map((value) => String(value || '').toLowerCase()).join(' ');
            return haystack.includes(term);
        })
        .sort((a, b) => {
            const emailCompare = normalizeEmail(a?.email).localeCompare(normalizeEmail(b?.email));
            if (emailCompare !== 0) return emailCompare;
            return String(a?.fullName || '').localeCompare(String(b?.fullName || ''));
        });
}

function renderScheduledReportAvailableUsers(searchTerm = '') {
    const select = document.getElementById('scheduledReportAvailableUsers');
    if (!select) return;
    const selectedRecipients = new Set(getScheduledReportRecipients());
    const users = getScheduledReportSelectableUsers(searchTerm);
    if (!users.length) {
        select.innerHTML = '<option value="">No matching users found</option>';
        return;
    }
    select.innerHTML = users.map((user) => {
        const email = normalizeEmail(user?.email);
        const fullName = String(user?.fullName || email.split('@')[0] || 'User').trim();
        const role = String(user?.role || 'user').trim() || 'user';
        const selectedTag = selectedRecipients.has(email) ? ' [Added]' : '';
        return `<option value="${escapeHtml(email)}">${escapeHtml(fullName)} - ${escapeHtml(email)} - ${escapeHtml(role)}${escapeHtml(selectedTag)}</option>`;
    }).join('');
}

function renderScheduledReportSelectedRecipients() {
    const host = document.getElementById('scheduledReportSelectedRecipients');
    if (!host) return;
    const recipients = getScheduledReportRecipients();
    if (!recipients.length) {
        host.innerHTML = '<div style="color:#94a3b8;font-size:13px;">No recipients selected yet.</div>';
        return;
    }
    host.innerHTML = recipients.map((email) => {
        const user = allUsers.find((entry) => normalizeEmail(entry?.email) === email);
        const label = String(user?.fullName || email).trim();
        return `
            <span style="display:inline-flex;align-items:center;gap:8px;padding:8px 10px;border-radius:999px;background:#eff6ff;border:1px solid #bfdbfe;color:#1e3a8a;font-size:13px;">
                <span><strong>${escapeHtml(label)}</strong><br><span style="font-size:11px;color:#64748b;">${escapeHtml(email)}</span></span>
                <button type="button" data-scheduled-report-remove="${escapeHtml(email)}" style="border:none;background:transparent;color:#b91c1c;font-weight:700;cursor:pointer;">x</button>
            </span>
        `;
    }).join('');
    host.querySelectorAll('[data-scheduled-report-remove]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const email = normalizeEmail(btn.getAttribute('data-scheduled-report-remove'));
            if (!email) return;
            removeScheduledReportRecipients([email]);
        });
    });
}

function refreshScheduledReportRecipientPicker() {
    const searchInput = document.getElementById('scheduledReportRecipientSearch');
    renderScheduledReportAvailableUsers(searchInput?.value || '');
    renderScheduledReportSelectedRecipients();
}

function addScheduledReportRecipients(emails = []) {
    const recipients = normalizeEmailList([...getScheduledReportRecipients(), ...emails]);
    syncScheduledReportRecipientsTextarea(recipients);
    refreshScheduledReportRecipientPicker();
}

function removeScheduledReportRecipients(emails = []) {
    const blocked = new Set(normalizeEmailList(emails));
    const recipients = getScheduledReportRecipients().filter((email) => !blocked.has(email));
    syncScheduledReportRecipientsTextarea(recipients);
    refreshScheduledReportRecipientPicker();
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
    if (normalized === 'reports_monitoring') return 'Reports Monitoring';
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
    const allowedRoles = new Set(['uploader', 'admin', 'super_admin', 'reviewer', 'rsa', 'payment', 'reports_monitoring']);
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

function formatUserLastLogin(u = {}) {
    return u.isOnline === true
        ? '<span class="status-badge approved">Online</span>'
        : escapeHtml(formatDate(u.lastLoginAt));
}

const reportDateKeyFormatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Africa/Lagos',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
});

function getLagosDateKey(date = new Date()) {
    return reportDateKeyFormatter.format(date);
}

function getPreviousScheduledReportDateKey(baseDate = new Date()) {
    const prior = new Date(baseDate);
    prior.setDate(prior.getDate() - 1);
    return getLagosDateKey(prior);
}

function getTimestampMillis(value) {
    if (!value) return 0;
    try {
        if (typeof value.toMillis === 'function') return value.toMillis();
        if (typeof value.seconds === 'number') return value.seconds * 1000;
        if (typeof value.toDate === 'function') return value.toDate().getTime();
    } catch (_) {}
    const parsed = new Date(value).getTime();
    return Number.isFinite(parsed) ? parsed : 0;
}

function getDateKey(value) {
    const ms = getTimestampMillis(value);
    if (!ms) return '';
    return getLagosDateKey(new Date(ms));
}

function isSameReportDate(value, dateKey) {
    return getDateKey(value) === String(dateKey || '').trim();
}

function isDateKeyWithinRange(value, startDateKey, endDateKey) {
    const dateKey = getDateKey(value);
    if (!dateKey) return false;
    const start = String(startDateKey || '').trim();
    const end = String(endDateKey || '').trim();
    if (!start || !end) return false;
    return dateKey >= start && dateKey <= end;
}

function parseMoneyValue(value) {
    const num = Number(String(value ?? '').replace(/[^0-9.\-]/g, ''));
    return Number.isFinite(num) ? num : 0;
}

function roundDownToThousand(value) {
    return Math.max(0, Math.floor(Number(value || 0) / 1000) * 1000);
}

function getSubmissionRsaBalance(sub = {}) {
    return parseMoneyValue(sub?.customerDetails?.rsaBalance || sub?.rsaBalance || 0);
}

function getSubmissionTwentyFivePercent(sub = {}) {
    const stored = parseMoneyValue(sub?.customerDetails?.rsa25Percent || sub?.rsa25Percent || 0);
    if (stored > 0) return roundDownToThousand(stored);
    return roundDownToThousand(getSubmissionRsaBalance(sub) * 0.25);
}

function getSubmissionCommissionOnePercent(sub = {}) {
    const base = getSubmissionTwentyFivePercent(sub);
    return Number((base * 0.01).toFixed(2));
}

function getUserDisplayNameByEmail(email = '') {
    const normalized = normalizeEmail(email);
    if (!normalized) return 'Unassigned';
    const user = allUsers.find((entry) => normalizeEmail(entry.email) === normalized);
    return String(user?.fullName || normalized).trim();
}

function getRejectionReason(sub = {}) {
    return String(sub?.latestRejectionReason || sub?.previousRejectionReason || sub?.comment || '').trim();
}

function getRejectionOfficerName(sub = {}) {
    if (getRejectionCount(sub) <= 0 && !getRejectionReason(sub)) return '';
    const rejectedBy = String(sub?.latestRejectedBy || sub?.previousRejectedBy || '').trim();
    if (rejectedBy) return getUserDisplayNameByEmail(rejectedBy);
    const stage = String(sub?.latestRejectedStage || '').trim().toLowerCase();
    if (stage === 'rsa') return getUserDisplayNameByEmail(sub?.assignedToRSA || '');
    if (stage === 'payment') return getUserDisplayNameByEmail(sub?.assignedToPayment || '');
    return getUserDisplayNameByEmail(sub?.reviewedBy || sub?.assignedTo || '');
}

function getRejectionCount(sub = {}) {
    const history = Array.isArray(sub?.rejectionHistory) ? sub.rejectionHistory.filter((entry) => {
        if (typeof entry === 'string') return String(entry).trim();
        return String(entry?.reason || '').trim();
    }) : [];
    if (history.length) return history.length;
    return getRejectionReason(sub) ? 1 : 0;
}

function formatMoneyForSheet(value) {
    const num = Number(value || 0);
    return num ? num : '';
}

function formatMoneyPreview(value) {
    const num = Number(value || 0);
    if (!Number.isFinite(num) || num === 0) return '-';
    return num.toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getPreviewStatus(sub = {}) {
    return String(sub.status || '').replace(/_/g, ' ');
}

function compareGroupedRows(a = {}, b = {}) {
    const ownerA = String(a.owner || '').toLowerCase();
    const ownerB = String(b.owner || '').toLowerCase();
    if (ownerA !== ownerB) return ownerA.localeCompare(ownerB);

    const assignedA = String(a.assignedAt || a.uploadedAt || '').toLowerCase();
    const assignedB = String(b.assignedAt || b.uploadedAt || '').toLowerCase();
    if (assignedA !== assignedB) return assignedA.localeCompare(assignedB);

    const stageA = String(a.stageTime || a.paidAt || a.clearedAt || '').toLowerCase();
    const stageB = String(b.stageTime || b.paidAt || b.clearedAt || '').toLowerCase();
    if (stageA !== stageB) return stageA.localeCompare(stageB);

    return String(a.customerName || '').toLowerCase().localeCompare(String(b.customerName || '').toLowerCase());
}

function buildUploaderSheetRows(records = []) {
    return records.map((sub) => ({
        owner: getUserDisplayNameByEmail(sub.uploadedBy),
        customerName: sub.customerName || '',
        rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
        rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
        commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
        status: String(sub.status || '').replace(/_/g, ' '),
        uploadedAt: formatDate(getSubmissionReviewEntryAt(sub)),
        rejectionReason: getRejectionReason(sub),
        rejectionOfficer: getRejectionOfficerName(sub),
        rejectionCount: getRejectionCount(sub)
    }));
}

function buildReviewerSheetRows(records = []) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedTo))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(sub.assignedTo),
            customerName: sub.customerName || '',
            uploaderName: getUserDisplayNameByEmail(sub.uploadedBy),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionReviewEntryAt(sub)),
            rejectionReason: getRejectionReason(sub),
            rejectionCount: getRejectionCount(sub)
        }));
}

function buildRsaSheetRows(records = []) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedToRSA))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(sub.assignedToRSA),
            customerName: sub.customerName || '',
            uploaderName: getUserDisplayNameByEmail(sub.uploadedBy),
            reviewerName: getUserDisplayNameByEmail(sub.assignedTo || sub.reviewedBy),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionRsaEntryAt(sub)),
            stageTime: formatDate(getSubmissionFinalSubmissionEntryAt(sub)),
            rejectionReason: getRejectionReason(sub),
            rejectionCount: getRejectionCount(sub)
        }));
}

function buildPaymentSheetRows(records = []) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedToPayment))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(sub.assignedToPayment),
            customerName: sub.customerName || '',
            uploaderName: getUserDisplayNameByEmail(sub.uploadedBy),
            rsaOfficerName: getUserDisplayNameByEmail(sub.assignedToRSA),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionPaymentEntryAt(sub)),
        }));
}

function normalizeOutstandingRowForStage(row = {}, stageId = '') {
    const normalizedStage = String(stageId || '').trim().toLowerCase();
    return row;
}

function normalizeRsaReportRow(row = {}) {
    const normalized = {
        owner: row.owner || '',
        customerName: row.customerName || '',
        uploaderName: row.uploaderName || '',
        reviewerName: row.reviewerName || '',
        rsaBalance: row.rsaBalance ?? '',
        rsa25: row.rsa25 ?? '',
        commission: row.commission ?? '',
        status: row.status || '',
        assignedAt: row.assignedAt || '-'
    };
    if (Object.prototype.hasOwnProperty.call(row, 'reportDate')) {
        return { owner: normalized.owner, reportDate: row.reportDate || '-', ...Object.fromEntries(Object.entries(normalized).filter(([key]) => key !== 'owner')) };
    }
    return normalized;
}

function applySheetHeaderStyle(row) {
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

function applySheetBodyStyle(row) {
    row.eachCell((cell) => {
        cell.font = { name: 'Calibri', size: 11 };
        cell.alignment = { vertical: 'middle', wrapText: true };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };
    });
}

function renderGroupedSheet({ worksheet, reportTitle, summaryRows = [], groupRows = [], tableHeaders = [], columns = [], decimalColumns = [], integerColumns = [] }) {
    worksheet.columns = columns;
    const titleRow = worksheet.addRow([reportTitle]);
    worksheet.mergeCells(`A${titleRow.number}:${String.fromCharCode(64 + tableHeaders.length)}${titleRow.number}`);
    titleRow.getCell(1).font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF0F3B67' } };

    summaryRows.forEach(([label, value]) => {
        const row = worksheet.addRow([label, value]);
        row.getCell(1).font = { bold: true };
    });

    worksheet.addRow([]);

    if (!groupRows.length) {
        worksheet.addRow(['No records found for the selected date.']);
        return;
    }

    let currentOwner = '';
    groupRows.forEach((item) => {
        if (item.owner !== currentOwner) {
            if (currentOwner) worksheet.addRow([]);
            currentOwner = item.owner;
            const ownerRow = worksheet.addRow([currentOwner]);
            worksheet.mergeCells(`A${ownerRow.number}:${String.fromCharCode(64 + tableHeaders.length)}${ownerRow.number}`);
            ownerRow.getCell(1).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF166534' } };
            ownerRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

            const headerRow = worksheet.addRow(tableHeaders);
            applySheetHeaderStyle(headerRow);
        }

        const dataRow = worksheet.addRow(Object.values(item).slice(1));
        applySheetBodyStyle(dataRow);
        dataRow.eachCell((cell, colNumber) => {
            if (decimalColumns.includes(colNumber) && typeof cell.value === 'number') {
                cell.numFmt = '#,##0.00';
            }
            if (integerColumns.includes(colNumber) && typeof cell.value === 'number') {
                cell.numFmt = '0';
            }
        });
    });
}

function buildDailyReportPreviewTable(sheet = {}) {
    const rows = Array.isArray(sheet.previewRows) ? sheet.previewRows : [];
    const headers = Array.isArray(sheet.previewHeaders) ? sheet.previewHeaders : [];
    if (!rows.length) {
        return '<div style="padding:18px;color:#64748b;">No records found for this sheet.</div>';
    }
    let currentOwner = '';
    const parts = [];
    rows.forEach((row) => {
        if (row.owner !== currentOwner) {
            if (currentOwner) parts.push('</tbody></table></div><div style="height:12px;"></div>');
            currentOwner = row.owner;
            parts.push(`
                <div style="padding:10px 12px;border:1px solid #dbe6f2;border-bottom:none;border-radius:12px 12px 0 0;background:#eefbf3;font-weight:700;color:#166534;">
                    ${escapeHtml(currentOwner)}
                </div>
                <div class="table-container" style="margin-bottom:0;">
                    <table class="documents-table">
                        <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
                        <tbody>
            `);
        }
        parts.push(`
            <tr>
                ${row.previewValues.map((value) => `<td>${escapeHtml(String(value ?? '-'))}</td>`).join('')}
            </tr>
        `);
    });
    parts.push('</tbody></table></div>');
    return parts.join('');
}

function renderScheduledReportPreviewTab(tabId) {
    if (!currentScheduledReportPreview) return;
    const tabsHost = document.getElementById('scheduledReportPreviewTabs');
    const summaryHost = document.getElementById('scheduledReportPreviewSummary');
    const contentHost = document.getElementById('scheduledReportPreviewContent');
    const titleEl = document.getElementById('scheduledReportPreviewTitle');
    const sheet = currentScheduledReportPreview.sheets.find((item) => item.id === tabId) || currentScheduledReportPreview.sheets[0];
    if (!sheet || !tabsHost || !summaryHost || !contentHost) return;

    currentScheduledReportPreview.activeTab = sheet.id;
    if (titleEl) titleEl.innerHTML = `<i class="fas fa-file-excel"></i> ${escapeHtml(sheet.title)}`;

    tabsHost.innerHTML = currentScheduledReportPreview.sheets.map((item) => `
        <button
            type="button"
            class="action-btn"
            data-report-tab="${escapeHtml(item.id)}"
            style="${item.id === sheet.id ? 'background:#003366;color:#fff;border:none;' : 'background:#e2e8f0;color:#0f172a;border:none;'}"
        >
            ${escapeHtml(item.tabLabel)}
        </button>
    `).join('');
    tabsHost.querySelectorAll('[data-report-tab]').forEach((button) => {
        button.addEventListener('click', () => renderScheduledReportPreviewTab(button.getAttribute('data-report-tab')));
    });

    summaryHost.innerHTML = sheet.summaryRows.map(([label, value]) => `
        <div style="padding:12px 14px;border:1px solid #dbe6f2;border-radius:12px;background:#f8fbff;">
            <div style="font-size:12px;color:#64748b;text-transform:uppercase;letter-spacing:.04em;">${escapeHtml(label)}</div>
            <div style="font-size:22px;font-weight:800;color:#0f3b67;margin-top:4px;">${escapeHtml(String(value))}</div>
        </div>
    `).join('');

    contentHost.innerHTML = buildDailyReportPreviewTable(sheet);
}

function openScheduledReportPreviewModal() {
    document.getElementById('scheduledReportPreviewModal')?.classList.add('active');
}

function closeScheduledReportPreviewModal() {
    document.getElementById('scheduledReportPreviewModal')?.classList.remove('active');
}

function buildDailyReportDefinition(reportDate) {
    const dateKey = String(reportDate || '').trim();
    if (!dateKey) {
        return null;
    }

    const submittedRecords = allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() !== 'draft');
    const reviewerOutstandingRecords = submittedRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending');
    const rsaOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return ['approved', 'processing_to_pfa'].includes(status) && !sub.finalSubmitted && !sub.rsaSubmitted;
    });
    const paymentOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return ['sent_to_pfa', 'rsa_submitted'].includes(status);
    });
    const uploaderRecords = submittedRecords.filter((sub) => isSameReportDate(getSubmissionReviewEntryAt(sub), dateKey));
    const reviewerRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedTo) && isSameReportDate(getSubmissionReviewEntryAt(sub), dateKey));
    const rsaRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedToRSA) && isSameReportDate(getSubmissionRsaEntryAt(sub), dateKey));
    const paymentRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedToPayment) && isSameReportDate(getSubmissionPaymentEntryAt(sub), dateKey));
    const reviewerAttendedRecords = reviewerRecords.filter((sub) => !!tsToMillis(getSubmissionApprovalEntryAt(sub)));
    const rsaAttendedRecords = rsaRecords.filter((sub) => !!tsToMillis(getSubmissionFinalSubmissionEntryAt(sub)) || String(sub.status || '').toLowerCase() === 'rejected_by_rsa');
    const paymentAttendedRecords = paymentRecords.filter((sub) => !!tsToMillis(getSubmissionPaidEntryAt(sub)) || !!tsToMillis(getSubmissionClearedEntryAt(sub)) || String(sub.status || '').toLowerCase() === 'cleared');

    const uploaderRows = buildUploaderSheetRows(uploaderRecords).sort(compareGroupedRows);
    const reviewerRows = buildReviewerSheetRows(reviewerRecords).sort(compareGroupedRows);
    const rsaRows = buildRsaSheetRows(rsaRecords).map(normalizeRsaReportRow).sort(compareGroupedRows);
    const paymentRows = buildPaymentSheetRows(paymentRecords).sort(compareGroupedRows);

    return {
        dateKey,
        activeTab: 'uploader',
        sheets: [
            {
                id: 'uploader',
                tabLabel: 'Uploader',
                title: `Uploader Report - ${dateKey}`,
                summaryRows: [
                    ['Total Uploaded', uploaderRecords.length],
                    ['Pending', uploaderRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length]
                ],
                rows: uploaderRows,
                excelHeaders: ['Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Uploaded Time', 'Reject Reason', 'Rejected By', 'Reject Count'],
                previewHeaders: ['Customer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Uploaded', 'Rejected By', 'Reject Count'],
                previewRows: uploaderRows.map((row) => ({
                    owner: row.owner,
                    previewValues: [row.customerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.uploadedAt, row.rejectionOfficer, String(row.rejectionCount || 0)]
                })),
                columns: [{ width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 24 }, { width: 14 }],
                decimalColumns: [2, 3, 4],
                integerColumns: [9]
            },
            {
                id: 'reviewer',
                tabLabel: 'Reviewer',
                title: `Reviewer Report - ${dateKey}`,
                summaryRows: [
                    ['Total Received', reviewerRecords.length],
                    ['Attended To', reviewerAttendedRecords.length],
                    ['Pending', reviewerRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length],
                    ['Total Outstanding', reviewerOutstandingRecords.length]
                ],
                rows: reviewerRows,
                excelHeaders: ['Customer Name', 'Uploader Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Reject Reason', 'Reject Count'],
                previewHeaders: ['Customer', 'Uploader', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned', 'Reject Count'],
                previewRows: reviewerRows.map((row) => ({
                    owner: row.owner,
                    previewValues: [row.customerName, row.uploaderName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt, String(row.rejectionCount || 0)]
                })),
                columns: [{ width: 28 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 14 }],
                decimalColumns: [3, 4, 5],
                integerColumns: [9]
            },
            {
                id: 'rsa',
                tabLabel: 'RSA',
                title: `RSA Report - ${dateKey}`,
                summaryRows: [
                    ['Total Received', rsaRecords.length],
                    ['Attended To', rsaAttendedRecords.length],
                    ['Pending', rsaRecords.filter((sub) => ['approved', 'processing_to_pfa'].includes(String(sub.status || '').toLowerCase()) && !sub.finalSubmitted && !sub.rsaSubmitted).length],
                    ['Total Outstanding', rsaOutstandingRecords.length]
                ],
                rows: rsaRows,
                excelHeaders: ['Customer Name', 'Uploader Name', 'Reviewer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'RSA Assigned Time'],
                previewHeaders: ['Customer', 'Uploader', 'Reviewer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned'],
                previewRows: rsaRows.map((row) => ({
                    owner: row.owner,
                    previewValues: [row.customerName, row.uploaderName, row.reviewerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt]
                })),
                columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
                decimalColumns: [4, 5, 6],
                integerColumns: []
            },
            {
                id: 'payment',
                tabLabel: 'Payment',
                title: `Payment Report - ${dateKey}`,
                summaryRows: [
                    ['Total Received', paymentRecords.length],
                    ['Attended To', paymentAttendedRecords.length],
                    ['Pending', paymentRecords.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase())).length],
                    ['Total Outstanding', paymentOutstandingRecords.length]
                ],
                rows: paymentRows,
                excelHeaders: ['Customer Name', 'Uploader Name', 'RSA Officer', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Payment Assigned Time'],
                previewHeaders: ['Customer', 'Uploader', 'RSA Officer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned'],
                previewRows: paymentRows.map((row) => ({
                    owner: row.owner,
                    previewValues: [row.customerName, row.uploaderName, row.rsaOfficerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt]
                })),
                columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
                decimalColumns: [4, 5, 6],
                integerColumns: []
            }
        ]
    };
}

function buildOutstandingReportDefinition({ reportDateKey = '', rangeStartDateKey = '', rangeEndDateKey = '', dashboard = 'all' } = {}) {
    const singleDateKey = String(reportDateKey || '').trim();
    const rangeStart = String(rangeStartDateKey || '').trim();
    const rangeEnd = String(rangeEndDateKey || '').trim();
    const hasRange = !!(rangeStart && rangeEnd);
    const startDateKey = hasRange ? (rangeStart <= rangeEnd ? rangeStart : rangeEnd) : '';
    const endDateKey = hasRange ? (rangeStart <= rangeEnd ? rangeEnd : rangeStart) : '';
    const dateKey = hasRange
        ? (startDateKey === REPORT_INCEPTION_START_DATE
            ? `${REPORT_INCEPTION_LABEL} to ${endDateKey}`
            : `${startDateKey} to ${endDateKey}`)
        : singleDateKey;
    if (!dateKey) return null;

    const submittedRecords = allSubmissions.filter((sub) => String(sub.status || '').toLowerCase() !== 'draft');
    const normalizedDashboard = String(dashboard || 'all').trim().toLowerCase() || 'all';
    const uploaderOutstandingRecords = submittedRecords.filter((sub) => (
        String(sub.status || '').toLowerCase() === 'pending'
    ));
    const reviewerOutstandingRecords = submittedRecords.filter((sub) => (
        String(sub.status || '').toLowerCase() === 'pending'
        && normalizeEmail(sub.assignedTo)
    ));
    const rsaOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return ['approved', 'processing_to_pfa'].includes(status)
            && !sub.finalSubmitted
            && !sub.rsaSubmitted
            && normalizeEmail(sub.assignedToRSA);
    });
    const paymentOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return ['sent_to_pfa', 'rsa_submitted'].includes(status)
            && normalizeEmail(sub.assignedToPayment);
    });

    const outstandingConfigs = [
        {
            id: 'uploader',
            tabLabel: 'Uploader',
            title: `Uploader Outstanding - ${dateKey}`,
            summaryRows: (records) => [
                ['Total Outstanding', records.length],
                ['Pending', records.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length]
            ],
            records: uploaderOutstandingRecords,
            rowsBuilder: buildUploaderSheetRows,
            excelHeaders: ['Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Uploaded Time', 'Reject Reason', 'Rejected By', 'Reject Count'],
            previewHeaders: ['Customer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Uploaded', 'Rejected By', 'Reject Count'],
            previewValues: (row) => [row.customerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.uploadedAt, row.rejectionOfficer, String(row.rejectionCount || 0)],
            columns: [{ width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 24 }, { width: 14 }],
            decimalColumns: [2, 3, 4],
            integerColumns: [9]
        },
        {
            id: 'reviewer',
            tabLabel: 'Reviewer',
            title: `Reviewer Outstanding - ${dateKey}`,
            summaryRows: (records) => [
                ['Total Outstanding', records.length],
                ['Pending', records.length]
            ],
            records: reviewerOutstandingRecords,
            rowsBuilder: buildReviewerSheetRows,
            excelHeaders: ['Customer Name', 'Uploader Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Reject Reason', 'Reject Count'],
            previewHeaders: ['Customer', 'Uploader', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned', 'Reject Count'],
            previewValues: (row) => [row.customerName, row.uploaderName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt, String(row.rejectionCount || 0)],
            columns: [{ width: 28 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 14 }],
            decimalColumns: [3, 4, 5],
            integerColumns: [9]
        },
        {
            id: 'rsa',
            tabLabel: 'RSA',
            title: `RSA Outstanding - ${dateKey}`,
            summaryRows: (records) => [
                ['Total Outstanding', records.length],
                ['Pending', records.length]
            ],
            records: rsaOutstandingRecords,
            rowsBuilder: buildRsaSheetRows,
            excelHeaders: ['Customer Name', 'Uploader Name', 'Reviewer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'RSA Assigned Time'],
            previewHeaders: ['Customer', 'Uploader', 'Reviewer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned'],
            previewValues: (row) => [row.customerName, row.uploaderName, row.reviewerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt],
            columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
            decimalColumns: [4, 5, 6],
            integerColumns: []
        },
        {
            id: 'payment',
            tabLabel: 'Payment',
            title: `Payment Outstanding - ${dateKey}`,
            summaryRows: (records) => [
                ['Total Outstanding', records.length],
                ['Pending', records.length]
            ],
            records: paymentOutstandingRecords,
            rowsBuilder: buildPaymentSheetRows,
            excelHeaders: ['Customer Name', 'Uploader Name', 'RSA Officer', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Payment Assigned Time'],
            previewHeaders: ['Customer', 'Uploader', 'RSA Officer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned'],
            previewValues: (row) => [row.customerName, row.uploaderName, row.rsaOfficerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt],
            columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
            decimalColumns: [4, 5, 6],
            integerColumns: []
        }
    ];

    const selectedConfigs = normalizedDashboard === 'all'
        ? outstandingConfigs
        : outstandingConfigs.filter((config) => config.id === normalizedDashboard);
    if (!selectedConfigs.length) return null;

    const sheets = selectedConfigs.map((config) => {
        const rows = config.rowsBuilder(config.records)
            .map((row) => normalizeOutstandingRowForStage(row, config.id))
            .map((row) => (config.id === 'rsa' ? normalizeRsaReportRow(row) : row))
            .sort(compareGroupedRows);
        return {
            id: config.id,
            tabLabel: config.tabLabel,
            title: config.title,
            summaryRows: config.summaryRows(config.records),
            rows,
            excelHeaders: config.excelHeaders,
            previewHeaders: config.previewHeaders,
            previewRows: rows.map((row) => ({
                owner: row.owner,
                previewValues: config.previewValues(row)
            })),
            columns: config.columns,
            decimalColumns: config.decimalColumns,
            integerColumns: config.integerColumns
        };
    });

    return {
        dateKey: hasRange ? startDateKey : singleDateKey,
        reportLabel: dateKey,
        activeTab: sheets[0]?.id || 'uploader',
        type: 'outstanding',
        selectedDashboard: normalizedDashboard,
        rangeStartDateKey: startDateKey,
        rangeEndDateKey: endDateKey,
        sheets
    };
}

async function downloadDailyReportWorkbook(reportInput) {
    const artifact = await buildDailyReportWorkbookArtifact(reportInput);
    if (!artifact) return;
    const { buffer, fileName, isOutstanding } = artifact;
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
    showNotification(isOutstanding ? 'Outstanding report Excel downloaded.' : 'Daily report Excel downloaded.', 'success');
}

async function buildDailyReportWorkbookArtifact(reportInput) {
    if (!window.ExcelJS) {
        showNotification('Excel export library is not available right now.', 'error');
        return null;
    }

    const report = typeof reportInput === 'string' ? buildDailyReportDefinition(reportInput) : reportInput;
    if (!report) {
        showNotification('Choose a report date first.', 'warning');
        return null;
    }

    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Super Admin';
    workbook.created = new Date();

    report.sheets.forEach((sheet) => {
        renderGroupedSheet({
            worksheet: workbook.addWorksheet(sheet.tabLabel + ' Report'),
            reportTitle: sheet.title,
            summaryRows: sheet.summaryRows,
            groupRows: sheet.rows,
            tableHeaders: sheet.excelHeaders,
            columns: sheet.columns,
            decimalColumns: sheet.decimalColumns,
            integerColumns: sheet.integerColumns
        });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const isOutstanding = report.type === 'outstanding';
    const dashboardSuffix = isOutstanding && report.selectedDashboard && report.selectedDashboard !== 'all'
        ? `_${report.selectedDashboard}`
        : '';
    const reportSuffix = String(report.reportLabel || report.dateKey || '')
        .replace(/\s+to\s+/gi, '_to_')
        .replace(/[^\w-]/g, '_');
    const fileName = isOutstanding
        ? `cmbank_outstanding_report${dashboardSuffix}_${reportSuffix}.xlsx`
        : `cmbank_daily_report_${report.dateKey}.xlsx`;

    return { report, buffer, fileName, isOutstanding };
}

function roleHome(role) {
    const r = String(role || '').toLowerCase();
    if (r === 'super_admin') return 'super-admin-dashboard.html';
    if (r === 'admin') return 'admin-dashboard.html';
    if (r === 'reports_monitoring') return 'reports-monitoring-dashboard.html';
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
    const onlineFilter = String(document.getElementById('superUserOnlineFilter')?.value || '').trim().toLowerCase();

    const filteredUsers = allUsers
        .filter((u) => {
            const role = String(u.role || 'uploader').toLowerCase();
            const status = String(u.status || 'active').toLowerCase();
            const presence = u.isOnline === true ? 'online' : 'offline';
            const whatsapp = getUserWhatsApp(u);
            const searchable = [
                u.fullName || '',
                u.displayName || '',
                u.email || '',
                whatsapp,
                getRoleLabel(role),
                status,
                presence
            ].join(' ').toLowerCase();
            return (!search || searchable.includes(search))
                && (!roleFilter || role === roleFilter)
                && (!statusFilter || status === statusFilter)
                && (!onlineFilter || presence === onlineFilter);
        })
        .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));

    if (!filteredUsers.length) {
        body.innerHTML = '<tr><td colspan="8" class="no-data">No users found</td></tr>';
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
                        <option value="reports_monitoring" ${role === 'reports_monitoring' ? 'selected' : ''}>Reports Monitoring</option>
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
                <td>${formatUserLastLogin(u)}</td>
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
    const privileged = allUsers.filter((u) => ['super_admin', 'admin', 'reports_monitoring', 'reviewer', 'rsa', 'payment'].includes(String(u.role || '').toLowerCase()));
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
    const scheduledReportEmailEnabled = document.getElementById('settingScheduledReportEmailEnabled');
    const scheduledReportSendTime = document.getElementById('settingScheduledReportSendTime');
    const scheduledReportSubject = document.getElementById('settingScheduledReportSubject');
    const scheduledReportRecipients = document.getElementById('settingScheduledReportRecipients');
    const scheduledReportBody = document.getElementById('settingScheduledReportBody');
    const scheduledReportDownloadDate = document.getElementById('scheduledReportDownloadDate');
    const scheduledOutstandingDownloadDate = document.getElementById('scheduledOutstandingDownloadDate');
    const scheduledOutstandingDownloadRangeStartDate = document.getElementById('scheduledOutstandingDownloadRangeStartDate');
    const scheduledOutstandingDownloadRangeEndDate = document.getElementById('scheduledOutstandingDownloadRangeEndDate');
    const scheduledOutstandingDownloadInceptionEndDate = document.getElementById('scheduledOutstandingDownloadInceptionEndDate');
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
    const forceLogoutCountdown = document.getElementById('settingForceLogoutCountdown');
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
    if (scheduledReportEmailEnabled) scheduledReportEmailEnabled.value = String(systemSettings.scheduledReportEmail?.enabled ? 'true' : 'false');
    if (scheduledReportSendTime) scheduledReportSendTime.value = String(systemSettings.scheduledReportEmail?.sendTime || defaultSystemSettings.scheduledReportEmail.sendTime || '08:00');
    if (scheduledReportSubject) scheduledReportSubject.value = String(systemSettings.scheduledReportEmail?.subject || defaultSystemSettings.scheduledReportEmail.subject || '');
    if (scheduledReportRecipients) scheduledReportRecipients.value = Array.isArray(systemSettings.scheduledReportEmail?.recipients) ? systemSettings.scheduledReportEmail.recipients.join('\n') : '';
    if (scheduledReportBody) scheduledReportBody.value = String(systemSettings.scheduledReportEmail?.body || defaultSystemSettings.scheduledReportEmail.body || '');
    if (scheduledReportDownloadDate && !scheduledReportDownloadDate.value) scheduledReportDownloadDate.value = getPreviousScheduledReportDateKey();
    if (scheduledOutstandingDownloadDate && !scheduledOutstandingDownloadDate.value) scheduledOutstandingDownloadDate.value = getPreviousScheduledReportDateKey();
    if (scheduledOutstandingDownloadRangeStartDate && !scheduledOutstandingDownloadRangeStartDate.value) scheduledOutstandingDownloadRangeStartDate.value = getPreviousScheduledReportDateKey();
    if (scheduledOutstandingDownloadRangeEndDate && !scheduledOutstandingDownloadRangeEndDate.value) scheduledOutstandingDownloadRangeEndDate.value = getPreviousScheduledReportDateKey();
    if (scheduledOutstandingDownloadInceptionEndDate && !scheduledOutstandingDownloadInceptionEndDate.value) scheduledOutstandingDownloadInceptionEndDate.value = getPreviousScheduledReportDateKey();
    renderScheduledReportDownloadTabState();
    updateOutstandingReportDownloadModeVisibility();
    refreshScheduledReportRecipientPicker();
    if (announcementEnabled) announcementEnabled.value = String(systemSettings.dashboardAnnouncement.enabled ? 'true' : 'false');
    if (announcementTone) announcementTone.value = String(systemSettings.dashboardAnnouncement.tone || 'info');
    if (announcementMessage) announcementMessage.value = String(systemSettings.dashboardAnnouncement.message || '');
    currentPfaOptions = Array.isArray(systemSettings.pfaOptions) ? [...systemSettings.pfaOptions] : [];
    currentPfaAddresses = { ...(systemSettings.pfaAddresses || {}) };
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
    if (forceLogoutCountdown) forceLogoutCountdown.value = String(systemSettings.securityControls.forceLogoutCountdown || '11m');
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
    renderScheduledReportLogs(false).catch(() => {});
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
    if (settingsModalSourceDropdownTab) {
        activeSettingsDropdownTab = settingsModalSourceDropdownTab;
        settingsModalSourceDropdownTab = '';
        if (currentTab === 'settings') renderSettingsSubTabState();
    }
}

function openSettingsSectionModal(panelId) {
    const panel = document.querySelector(`[data-settings-panel-id="${panelId}"]`);
    if (!panel || !settingsSectionModalContentHost) return;

    returnActiveSettingsPanelToStorage();
    settingsModalSourceDropdownTab = activeSettingsDropdownTab || currentSettingsSubTab || '';
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
        body.innerHTML = '<tr><td colspan="5" class="no-data">No PFAs configured yet</td></tr>';
        return;
    }
    body.innerHTML = currentPfaOptions.map((name) => {
        const encodedName = encodeURIComponent(name);
        const addressInfo = normalizePfaAddressEntry(currentPfaAddresses?.[name]);
        return `
        <tr>
            <td><strong>${escapeHtml(name)}</strong></td>
            <td>
                <input
                    type="text"
                    placeholder="Address line"
                    value="${escapeHtml(addressInfo.address)}"
                    oninput="window.updatePfaAddressField('${encodedName}', 'address', this.value)"
                    style="width:100%;min-width:220px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"
                >
            </td>
            <td>
                <input
                    type="text"
                    placeholder="Landmark / City"
                    value="${escapeHtml(addressInfo.landmark)}"
                    oninput="window.updatePfaAddressField('${encodedName}', 'landmark', this.value)"
                    style="width:100%;min-width:170px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"
                >
            </td>
            <td>
                <input
                    type="text"
                    placeholder="State"
                    value="${escapeHtml(addressInfo.state)}"
                    oninput="window.updatePfaAddressField('${encodedName}', 'state', this.value)"
                    style="width:100%;min-width:130px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"
                >
            </td>
            <td><button class="action-btn" type="button" onclick="window.removePfaOption('${encodedName}')" style="background:#b91c1c;color:#fff;border:none;"><i class="fas fa-trash"></i> Remove</button></td>
        </tr>
    `;
    }).join('');
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
    currentPfaAddresses = { ...currentPfaAddresses, [name]: normalizePfaAddressEntry(currentPfaAddresses?.[name]) };
    if (input) input.value = '';
    renderPfaManagement();
};

function normalizePfaAddressEntry(value = {}) {
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

window.updatePfaAddressField = (encodedName, field, value) => {
    const name = decodeURIComponent(String(encodedName || ''));
    const current = normalizePfaAddressEntry(currentPfaAddresses?.[name]);
    currentPfaAddresses = {
        ...currentPfaAddresses,
        [name]: {
            ...current,
            [field]: String(value || '').trim()
        }
    };
};

window.removePfaOption = (encodedName) => {
    const name = decodeURIComponent(String(encodedName || ''));
    currentPfaOptions = currentPfaOptions.filter((item) => item !== name);
    const nextAddresses = { ...currentPfaAddresses };
    delete nextAddresses[name];
    currentPfaAddresses = nextAddresses;
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
            rsaAssignedAt: rsa ? serverTimestamp() : null,
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
    const scheduledReportEmailEnabled = String(document.getElementById('settingScheduledReportEmailEnabled')?.value || 'false') === 'true';
    const scheduledReportSendTime = String(document.getElementById('settingScheduledReportSendTime')?.value || getDefaultSystemSettings().scheduledReportEmail.sendTime || '08:00').trim();
    const scheduledReportSubject = String(document.getElementById('settingScheduledReportSubject')?.value || '').trim() || getDefaultSystemSettings().scheduledReportEmail.subject;
    const scheduledReportRecipients = normalizeEmailList(parseLinesTextarea('settingScheduledReportRecipients'));
    const scheduledReportBody = String(document.getElementById('settingScheduledReportBody')?.value || '').trim() || getDefaultSystemSettings().scheduledReportEmail.body;
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
    const forceLogoutCountdown = parseForceLogoutDurationInput(
        document.getElementById('settingForceLogoutCountdown')?.value || getDefaultSystemSettings().securityControls.forceLogoutCountdown || '11m'
    );
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
    if (!forceLogoutCountdown || forceLogoutDurationToSeconds(forceLogoutCountdown) <= 0) {
        showNotification('Enter a valid force logout countdown like 30s, 10m, or 1h', 'warning');
        return false;
    }
    if (!Number.isFinite(auditRetentionDays) || auditRetentionDays <= 0) {
        showNotification('Audit retention days must be greater than 0', 'warning');
        return false;
    }
    if (!/^\d{2}:\d{2}$/.test(scheduledReportSendTime)) {
        showNotification('Enter a valid report send time', 'warning');
        return false;
    }
    if (scheduledReportEmailEnabled && !scheduledReportRecipients.length) {
        showNotification('Add at least one recipient email for the scheduled report', 'warning');
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
    const pfaAddresses = Object.fromEntries(
        pfaOptions
            .map((name) => [name, normalizePfaAddressEntry(currentPfaAddresses?.[name])])
            .filter(([, entry]) => entry.address || entry.landmark || entry.state)
    );
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
            scheduledReportEmail: {
                enabled: scheduledReportEmailEnabled,
                sendTime: scheduledReportSendTime,
                subject: scheduledReportSubject,
                recipients: scheduledReportRecipients,
                body: scheduledReportBody,
                reportDateMode: 'previous_day'
            },
            pfaOptions,
            pfaAddresses,
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
                forceLogoutCountdown,
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
            scheduledReportEmailEnabled,
            scheduledReportSendTime,
            scheduledReportRecipientCount: scheduledReportRecipients.length,
            announcementEnabled,
            defaultRouteMode,
            agentRegistrationApprovalRequired,
            rejectionMinLength,
            pfaCount: pfaOptions.length,
            documentRequirementCount: Array.isArray(documentRequirements) ? documentRequirements.length : 0,
            documentRequirementRoles,
            bulkImportColumnCount: bulkImportRequiredColumns.length,
            sessionTimeoutMinutes,
            forceLogoutCountdown,
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
    try {
        const currentSettings = await getSystemSettings(db, { force: true });
        const countdownSetting = parseForceLogoutDurationInput(currentSettings.securityControls?.forceLogoutCountdown || getDefaultSystemSettings().securityControls.forceLogoutCountdown || '11m');
        const countdownLabel = describeForceLogoutDuration(countdownSetting);
        const confirmed = confirm(`Force logout all signed-in users?\n\nThey will first see a warning and can finish current work before automatic logout in ${countdownLabel}.`);
        if (!confirmed) return;

        const securityControls = {
            ...(currentSettings.securityControls || {}),
            forceLogoutToken: new Date().toISOString()
        };
        const refreshToken = `${Date.now()}`;
        await setDoc(doc(db, 'settings', 'system'), {
            securityControls,
            cacheClearToken: refreshToken,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        }, { merge: true });
        clearSystemSettingsCache();
        await addDoc(collection(db, 'audit'), {
            action: 'force_logout_all_users',
            countdown: countdownSetting,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification(`Force logout notice sent successfully. Open dashboards will refresh and automatic logout will happen in ${countdownLabel}.`, 'success');
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
    } catch (_) {}
    window.location.href = 'index.html';
};

function setupRealtimeData() {
    onSnapshot(query(collection(db, 'users')), (snap) => {
        allUsers = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        refreshScheduledReportRecipientPicker();
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
document.getElementById('superUserOnlineFilter')?.addEventListener('change', renderAdminManagement);
document.getElementById('routingRulesSearch')?.addEventListener('input', () => {
    if (currentTab === 'routing-rules') renderRoutingRules();
});
document.getElementById('refreshScheduledReportLogsInlineBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('refreshScheduledReportLogsInlineBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Refreshing...';
        }
        await renderScheduledReportLogs(true);
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('refreshScheduledReportLogsModalBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('refreshScheduledReportLogsModalBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Refreshing...';
        }
        await renderScheduledReportLogs(true);
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
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
document.getElementById('previewScheduledReportBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('previewScheduledReportBtn');
    const originalHtml = button?.innerHTML || '';
    const reportDateInput = document.getElementById('scheduledReportDownloadDate');
    let reportDate = String(reportDateInput?.value || '').trim();
    try {
        if (!reportDate) {
            reportDate = getPreviousScheduledReportDateKey();
            if (reportDateInput) reportDateInput.value = reportDate;
        }
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Preparing...';
        }
        const report = buildDailyReportDefinition(reportDate);
        if (!report) {
            showNotification('Choose a report date first.', 'warning');
            return;
        }
        currentScheduledReportPreview = report;
        renderScheduledReportPreviewTab(report.activeTab || report.sheets[0]?.id || 'uploader');
        openScheduledReportPreviewModal();
    } catch (_) {
        showNotification('Failed to prepare daily report preview.', 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('previewOutstandingReportBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('previewOutstandingReportBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Preparing...';
        }
        const selection = buildOutstandingDownloadSelection();
        const report = buildOutstandingReportDefinition(selection.reportOptions);
        if (!report) {
            showNotification('Choose a valid outstanding report date first.', 'warning');
            return;
        }
        currentScheduledReportPreview = report;
        renderScheduledReportPreviewTab(report.activeTab || report.sheets[0]?.id || 'uploader');
        openScheduledReportPreviewModal();
    } catch (error) {
        const message = String(error?.message || 'Failed to prepare outstanding report preview.');
        console.error('Outstanding report preview failed:', error);
        showNotification(message, 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('sendOutstandingReportBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('sendOutstandingReportBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        const selection = buildOutstandingDownloadSelection();
        const report = buildOutstandingReportDefinition(selection.reportOptions);
        if (!report) {
            showNotification('Choose a valid outstanding report date first.', 'warning');
            return;
        }
        setOutstandingReportDownloadStatus(`Preparing to send ${selection.label}...`, 'info');
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Sending...';
        }
        await sendPrebuiltScheduledWorkbook({
            report,
            reportLabel: selection.label,
            setStatus: setOutstandingReportDownloadStatus,
            successMessage: 'Outstanding report sent successfully.',
            busyMessage: `Sending ${selection.label} to configured scheduled report recipients...`
        });
        showNotification('Outstanding report sent successfully.', 'success');
    } catch (error) {
        const message = String(error?.name === 'AbortError' ? 'Timed out while sending the outstanding report.' : (error?.message || 'Failed to send outstanding report.'));
        setOutstandingReportDownloadStatus(message, 'error');
        showNotification(message, 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('downloadScheduledReportFromModalBtn')?.addEventListener('click', async () => {
    try {
        await downloadDailyReportWorkbook(currentScheduledReportPreview);
    } catch (_) {
        showNotification('Failed to download daily report workbook.', 'error');
    }
});
document.getElementById('closeScheduledReportPreviewModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportPreviewModal();
});
document.getElementById('cancelScheduledReportPreviewModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportPreviewModal();
});
document.getElementById('scheduledReportPreviewModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'scheduledReportPreviewModal') {
        event.stopPropagation();
        closeScheduledReportPreviewModal();
    }
});
document.getElementById('scheduledReportRecipientSearch')?.addEventListener('input', (event) => {
    renderScheduledReportAvailableUsers(event.target?.value || '');
});
document.getElementById('addScheduledReportRecipientsBtn')?.addEventListener('click', () => {
    const select = document.getElementById('scheduledReportAvailableUsers');
    if (!select) return;
    const selectedEmails = Array.from(select.selectedOptions || [])
        .map((option) => normalizeEmail(option.value))
        .filter(Boolean);
    if (!selectedEmails.length) {
        showNotification('Select at least one user email to add.', 'warning');
        return;
    }
    addScheduledReportRecipients(selectedEmails);
});
document.getElementById('addAllScheduledReportRecipientsBtn')?.addEventListener('click', () => {
    const searchInput = document.getElementById('scheduledReportRecipientSearch');
    const emails = getScheduledReportSelectableUsers(searchInput?.value || '')
        .map((user) => normalizeEmail(user?.email))
        .filter(Boolean);
    if (!emails.length) {
        showNotification('No user emails available to add.', 'warning');
        return;
    }
    addScheduledReportRecipients(emails);
});
document.getElementById('settingScheduledReportRecipients')?.addEventListener('input', () => {
    refreshScheduledReportRecipientPicker();
});
document.getElementById('sendCustomScheduledReportBtn')?.addEventListener('click', () => {
    openScheduledReportSendModal();
});
document.getElementById('scheduledReportSendDailyTabBtn')?.addEventListener('click', () => {
    currentScheduledReportSendTab = 'daily';
    renderScheduledReportSendTabState();
});
document.getElementById('scheduledReportSendOutstandingTabBtn')?.addEventListener('click', () => {
    currentScheduledReportSendTab = 'outstanding';
    renderScheduledReportSendTabState();
});
document.getElementById('scheduledReportDownloadDailyTabBtn')?.addEventListener('click', () => {
    currentScheduledReportDownloadTab = 'daily';
    renderScheduledReportDownloadTabState();
});
document.getElementById('scheduledReportDownloadOutstandingTabBtn')?.addEventListener('click', () => {
    currentScheduledReportDownloadTab = 'outstanding';
    renderScheduledReportDownloadTabState();
});
document.getElementById('openScheduledReportLogsModalBtn')?.addEventListener('click', () => {
    openScheduledReportLogsModal();
});
document.getElementById('scheduledReportSendMode')?.addEventListener('change', () => {
    updateScheduledReportSendModeVisibility();
});
document.getElementById('scheduledOutstandingSendMode')?.addEventListener('change', () => {
    updateOutstandingReportSendModeVisibility();
});
document.getElementById('scheduledOutstandingDownloadMode')?.addEventListener('change', () => {
    updateOutstandingReportDownloadModeVisibility();
});
document.getElementById('closeScheduledReportSendModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportSendModal();
});
document.getElementById('cancelScheduledReportSendModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportSendModal();
});
document.getElementById('scheduledReportSendModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'scheduledReportSendModal') {
        event.stopPropagation();
        closeScheduledReportSendModal();
    }
});
document.getElementById('closeScheduledReportLogsModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportLogsModal();
});
document.getElementById('cancelScheduledReportLogsModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportLogsModal();
});
document.getElementById('scheduledReportLogsModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'scheduledReportLogsModal') {
        event.stopPropagation();
        closeScheduledReportLogsModal();
    }
});
document.getElementById('closeScheduledReportConfirmModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportConfirmModal(false);
});
document.getElementById('cancelScheduledReportConfirmModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportConfirmModal(false);
});
document.getElementById('confirmScheduledReportConfirmModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeScheduledReportConfirmModal(true);
});
document.getElementById('scheduledReportConfirmModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'scheduledReportConfirmModal') {
        event.stopPropagation();
        closeScheduledReportConfirmModal(false);
    }
});
document.getElementById('checkScheduledReportStatusBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('checkScheduledReportStatusBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        setScheduledReportSendStatus('Checking scheduled report sender...', 'info');
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Checking...';
        }
        const status = await getScheduledReportBackendStatus();
        if (status.enabled !== true) {
            setScheduledReportSendStatus('Scheduled report sender is reachable, but daily report email is disabled in settings.', 'warning', describeScheduledReportStatus(status), status);
            return;
        }
        if (!Array.isArray(status.recipients) || !status.recipients.length) {
            setScheduledReportSendStatus('Scheduled report sender is reachable, but no recipients are configured.', 'warning', describeScheduledReportStatus(status), status);
            return;
        }
        setScheduledReportSendStatus('Scheduled report sender is connected and ready.', 'success', describeScheduledReportStatus(status), status);
    } catch (error) {
        const message = String(error?.name === 'AbortError' ? 'Timed out while checking sender status.' : (error?.message || 'Failed to check sender status.'));
        setScheduledReportSendStatus(message, 'error');
        showNotification(message, 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('confirmScheduledReportSendBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('confirmScheduledReportSendBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        const selection = buildManualScheduledReportPayload();
        setScheduledReportSendModalStatus('Running checks before sending the selected report...', 'info');
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Sending...';
        }
        await sendSelectedScheduledReport(selection, {
            setStatus: setScheduledReportSendModalStatus,
            successMessage: 'Selected report sent successfully.'
        });
        showNotification('Custom report email sent successfully.', 'success');
    } catch (error) {
        const message = String(error?.name === 'AbortError' ? 'Timed out while sending the selected report.' : (error?.message || 'Failed to send selected report.'));
        setScheduledReportSendModalStatus(message, 'error');
        showNotification(message, 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('sendScheduledReportNowBtn')?.addEventListener('click', async () => {
    const button = document.getElementById('sendScheduledReportNowBtn');
    const originalHtml = button?.innerHTML || '';
    try {
        setScheduledReportSendStatus('Running preflight checks before sending previous day report...', 'info');
        const status = await getScheduledReportBackendStatus();
        if (status.enabled !== true) {
            const message = 'Scheduled report email is disabled in settings.';
            setScheduledReportSendStatus(message, 'warning', describeScheduledReportStatus(status), status);
            showNotification(message, 'warning');
            return;
        }
        if (!Array.isArray(status.recipients) || !status.recipients.length) {
            const message = 'No scheduled report recipients are configured.';
            setScheduledReportSendStatus(message, 'warning', describeScheduledReportStatus(status), status);
            showNotification(message, 'warning');
            return;
        }
        const apiBaseUrl = getEmailApiBaseUrl();
        if (!apiBaseUrl) {
            throw new Error('Scheduled report sender backend URL is not configured.');
        }
        if (!currentUser) {
            throw new Error('You must be signed in to trigger the scheduled report.');
        }
        if (button) {
            button.disabled = true;
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Sending...';
        }
        setScheduledReportSendStatus('Sending previous day report now...', 'info', describeScheduledReportStatus(status), status);
        const idToken = await currentUser.getIdToken(true);
        let forceResend = false;
        let requestResult = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/send-now`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${idToken}`
            },
            body: JSON.stringify({})
        }, 60000);
        if (requestResult.response.status === 409 || String(requestResult.data?.error || '').trim() === 'already-sent') {
            if (button) {
                button.disabled = false;
                button.innerHTML = originalHtml;
            }
            const confirmed = await openScheduledReportConfirmModal({
                title: 'Report Already Sent',
                heroTitle: 'Previous day report was already sent',
                message: 'Previous day report has already been sent. Do you still want to resend it?',
                note: 'Continuing will resend the previous day report to the configured recipient list.',
                confirmLabel: 'Resend Report'
            });
            if (!confirmed) {
                setScheduledReportSendStatus('Manual resend cancelled. The previous day report had already been sent earlier.', 'info', describeScheduledReportStatus(status), status);
                showNotification('Scheduled report resend cancelled.', 'info');
                return;
            }
            forceResend = true;
            if (button) {
                button.disabled = true;
                button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Resending...';
            }
            setScheduledReportSendStatus('Resending previous day report on your confirmation...', 'warning', describeScheduledReportStatus(status), status);
            requestResult = await fetchJsonWithTimeout(`${apiBaseUrl}/api/scheduled-report/send-now`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${idToken}`
                },
                body: JSON.stringify({ forceResend: true })
            }, 60000);
        }
        const { response, data: result } = requestResult;
        if (!response.ok || result?.ok === false) {
            throw new Error(String(result?.error || result?.reason || 'Failed to send scheduled report'));
        }
        const sentCount = Number(result?.sentCount || 0);
        const failedCount = Number(result?.failedCount || 0);
        const detailLines = [
            `Report date sent: ${String(result?.reportDateKey || status.reportDateKey || '-').trim() || '-'}`,
            `Excel file: ${String(result?.attachmentFileName || '-').trim() || '-'}`,
            `Successful recipients: ${sentCount}`,
            `Failed recipients: ${failedCount}`
        ];
        if (Array.isArray(result?.failures)) {
            result.failures.slice(0, 5).forEach((entry) => {
                detailLines.push(`${String(entry?.recipient || 'Recipient')}: ${String(entry?.error || 'failed')}`);
            });
        }
        const refreshedStatus = await getScheduledReportBackendStatus().catch(() => null);
        setScheduledReportSendStatus(
            failedCount > 0
                ? (forceResend ? 'Scheduled report was resent, but some recipients failed.' : 'Scheduled report was sent, but some recipients failed.')
                : (forceResend ? 'Scheduled report resent successfully.' : 'Scheduled report sent successfully.'),
            failedCount > 0 ? 'warning' : 'success',
            detailLines,
            refreshedStatus
        );
        showNotification(
            failedCount > 0
                ? `Scheduled report sent with ${sentCount} success and ${failedCount} failure(s).`
                : (forceResend
                    ? `Scheduled report resent successfully to ${sentCount} recipient(s).`
                    : `Scheduled report sent successfully to ${sentCount} recipient(s).`),
            failedCount > 0 ? 'warning' : 'success'
        );
    } catch (error) {
        const message = String(error?.name === 'AbortError' ? 'Timed out while sending scheduled report.' : (error?.message || 'Failed to send scheduled report'));
        setScheduledReportSendStatus(message, 'error');
        showNotification(message, 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.innerHTML = originalHtml;
        }
    }
});
document.getElementById('saveSettingsBtn')?.addEventListener('click', () => {
    window.saveSuperSettings?.();
});
document.getElementById('saveSettingsSectionModalBtn')?.addEventListener('click', async (event) => {
    const ok = await window.saveSuperSettings?.(event.currentTarget);
    if (ok) closeSettingsSectionModal();
});
document.getElementById('closeSettingsSectionModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeSettingsSectionModal();
});
document.getElementById('cancelSettingsSectionModalBtn')?.addEventListener('click', (event) => {
    event.stopPropagation();
    closeSettingsSectionModal();
});
document.getElementById('settingsSectionModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'settingsSectionModal') {
        event.stopPropagation();
        closeSettingsSectionModal();
    }
});
document.addEventListener('click', (event) => {
    if (!settingsDropdownNav) return;
    if (settingsSectionModal?.classList.contains('active')) return;
    if (!settingsDropdownNav.contains(event.target)) {
        if (activeSettingsDropdownTab) {
            activeSettingsDropdownTab = '';
            renderSettingsSubTabState();
        }
    }
});
document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && document.getElementById('scheduledReportConfirmModal')?.classList.contains('active')) {
        closeScheduledReportConfirmModal(false);
        return;
    }
    if (event.key === 'Escape' && document.getElementById('scheduledReportSendModal')?.classList.contains('active')) {
        closeScheduledReportSendModal();
        return;
    }
    if (event.key === 'Escape' && document.getElementById('scheduledReportLogsModal')?.classList.contains('active')) {
        closeScheduledReportLogsModal();
        return;
    }
    if (event.key === 'Escape' && document.getElementById('scheduledReportPreviewModal')?.classList.contains('active')) {
        closeScheduledReportPreviewModal();
        return;
    }
    if (event.key === 'Escape' && settingsSectionModal?.classList.contains('active')) {
        closeSettingsSectionModal();
        return;
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
