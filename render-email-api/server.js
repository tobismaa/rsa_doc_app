require('dotenv').config();

const express = require('express');
const cors = require('cors');
const admin = require('firebase-admin');
const sgMail = require('@sendgrid/mail');
const { createDailyReportWorkbookBuffer, getPreviousDateKeyInLagos, getLagosDateKey } = require('./scheduled-report');

const app = express();
const port = Number(process.env.PORT || 3000);
const requireAuth = String(process.env.REQUIRE_FIREBASE_AUTH || 'true').toLowerCase() !== 'false';
const allowedOrigins = String(process.env.ALLOWED_ORIGINS || '')
    .split(',')
    .map((v) => v.trim())
    .filter(Boolean);

const allowedTypes = new Set([
    'submission_assigned_viewer',
    'submission_rejected_uploader',
    'submission_approved_rsa',
    'submission_approved_uploader',
    'submission_final_uploader'
]);
const EMAILJS_PUBLIC_KEY = String(process.env.EMAILJS_PUBLIC_KEY || '').trim();
const EMAILJS_SERVICE_ID = String(process.env.EMAILJS_SERVICE_ID || '').trim();
const EMAILJS_TEMPLATE_ID = String(process.env.EMAILJS_TEMPLATE_ID || '').trim();
const EMAILJS_REPORT_TEMPLATE_ID = String(process.env.EMAILJS_REPORT_TEMPLATE_ID || EMAILJS_TEMPLATE_ID || '').trim();
const EMAILJS_PRIVATE_KEY = String(process.env.EMAILJS_PRIVATE_KEY || '').trim();
const EMAILJS_ATTACHMENT_PARAM = String(process.env.EMAILJS_ATTACHMENT_PARAM || 'report_attachment').trim() || 'report_attachment';
const EMAILJS_ATTACHMENT_FILENAME_PARAM = String(process.env.EMAILJS_ATTACHMENT_FILENAME_PARAM || 'report_attachment_filename').trim() || 'report_attachment_filename';
const ENABLE_SCHEDULED_REPORT_SENDER = String(process.env.ENABLE_SCHEDULED_REPORT_SENDER || 'true').toLowerCase() !== 'false';

function initFirebaseAdmin() {
    if (admin.apps.length > 0) return;

    const raw = process.env.FIREBASE_SERVICE_ACCOUNT_JSON;
    if (raw) {
        const serviceAccount = JSON.parse(raw);
        admin.initializeApp({
            credential: admin.credential.cert(serviceAccount)
        });
        return;
    }

    admin.initializeApp();
}

function normalizeEmail(value) {
    const v = String(value || '').trim().toLowerCase();
    return v.includes('@') ? v : '';
}

function normalizeRole(value) {
    return String(value || '').trim().toLowerCase();
}

function safeLowerEmail(value) {
    return normalizeEmail(value);
}

function escapeHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function normalizeEmailList(values = []) {
    const seen = new Set();
    const out = [];
    values.forEach((value) => {
        const email = normalizeEmail(value);
        if (!email || seen.has(email)) return;
        seen.add(email);
        out.push(email);
    });
    return out;
}

function uniqueEmails(values = []) {
    return Array.from(new Set(values.map((value) => safeLowerEmail(value)).filter(Boolean)));
}

function getSubmissionParticipantEmails(sub = {}) {
    return uniqueEmails([
        sub.uploadedBy,
        sub.assignedTo,
        sub.assignedToRSA,
        sub.assignedToPayment
    ]);
}

function getStatusPushRecipientEmails(sub = {}, newStatus = '') {
    const status = String(newStatus || sub.status || '').trim().toLowerCase();
    const uploader = safeLowerEmail(sub.uploadedBy);
    const reviewer = safeLowerEmail(sub.assignedTo);
    const rsa = safeLowerEmail(sub.assignedToRSA);
    const payment = safeLowerEmail(sub.assignedToPayment);

    if (['pending', 'submitted', 'resubmitted'].includes(status)) {
        return uniqueEmails([reviewer]);
    }

    if (['rejected', 'rejected_by_reviewer', 'rejected_by_rsa'].includes(status)) {
        return uniqueEmails([uploader]);
    }

    if (['approved', 'processing_to_pfa'].includes(status)) {
        return uniqueEmails([uploader, rsa]);
    }

    if (['sent_to_pfa', 'rsa_submitted'].includes(status)) {
        return uniqueEmails([uploader]);
    }

    if (['paid', 'cleared'].includes(status)) {
        return uniqueEmails([uploader]);
    }

    return uniqueEmails([uploader]);
}

function getSubmissionStageKey(sub = {}) {
    const status = String(sub.status || '').trim().toLowerCase();
    if (status === 'cleared') return 'closed';
    if (['sent_to_pfa', 'rsa_submitted', 'paid'].includes(status)) return 'payment';
    if (['processing_to_pfa', 'approved'].includes(status)) return 'rsa';
    return 'review';
}

function getStageHandlerEmail(sub = {}, stageKey = '') {
    const stage = String(stageKey || '').trim().toLowerCase();
    if (stage === 'review') return safeLowerEmail(sub.assignedTo || sub.reviewedBy);
    if (stage === 'rsa') return safeLowerEmail(sub.assignedToRSA);
    if (stage === 'payment') return safeLowerEmail(sub.assignedToPayment);
    return '';
}

function getChatPushRecipientEmails(sub = {}, chatMeta = {}) {
    const stageKey = getSubmissionStageKey(sub);
    if (stageKey === 'closed') return [];

    const recipients = [
        safeLowerEmail(sub.uploadedBy),
        getStageHandlerEmail(sub, stageKey)
    ];

    if (chatMeta?.escalated === true) {
        (Array.isArray(chatMeta.adminParticipants) ? chatMeta.adminParticipants : [])
            .forEach((email) => recipients.push(email));
    }

    return uniqueEmails(recipients);
}

function resolveAllowedOrigin(req) {
    const reqOrigin = String(req.header('Origin') || '').trim();
    if (reqOrigin && (allowedOrigins.length === 0 || allowedOrigins.includes(reqOrigin))) {
        return reqOrigin;
    }
    return allowedOrigins[0] || '';
}

function buildResetPageUrl(req, candidateUrl) {
    const raw = String(candidateUrl || '').trim();
    const fallbackOrigin = resolveAllowedOrigin(req);
    const fallbackUrl = fallbackOrigin ? `${fallbackOrigin.replace(/\/+$/, '')}/reset-password.html` : '';

    if (!raw) return fallbackUrl;

    try {
        const parsed = new URL(raw);
        const parsedOrigin = parsed.origin;
        const allowed = allowedOrigins.length === 0 || allowedOrigins.includes(parsedOrigin);
        return allowed ? parsed.toString() : fallbackUrl;
    } catch (_) {
        return fallbackUrl;
    }
}

const lagosDateTimeFormatter = new Intl.DateTimeFormat('en-NG', {
    timeZone: 'Africa/Lagos',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
});

const lagosDateKeyFormatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Africa/Lagos',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
});

const lagosTimeFormatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Africa/Lagos',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
});

const lagosWeekdayFormatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Africa/Lagos',
    weekday: 'short'
});

function getLagosTimeKey(date = new Date()) {
    return lagosTimeFormatter.format(date);
}

function getLagosWeekdayKey(date = new Date()) {
    return lagosWeekdayFormatter.format(date).toLowerCase();
}

function tsToMillis(value) {
    if (!value) return 0;
    try {
        if (typeof value.toMillis === 'function') return value.toMillis();
        if (typeof value.toDate === 'function') return value.toDate().getTime();
        if (typeof value.seconds === 'number') return value.seconds * 1000;
    } catch (_) {}
    const parsed = new Date(value).getTime();
    return Number.isFinite(parsed) ? parsed : 0;
}

function formatLagosDateTime(value) {
    const ms = tsToMillis(value);
    if (!ms) return '';
    return lagosDateTimeFormatter.format(new Date(ms));
}

function getLagosDateKeyFromValue(value) {
    const ms = tsToMillis(value);
    if (!ms) return '';
    return getLagosDateKey(new Date(ms));
}

function shiftDateKey(dateKey, offsetDays) {
    const text = String(dateKey || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return '';
    const [year, month, day] = text.split('-').map(Number);
    const utcDate = new Date(Date.UTC(year, month - 1, day));
    utcDate.setUTCDate(utcDate.getUTCDate() + Number(offsetDays || 0));
    return utcDate.toISOString().slice(0, 10);
}

function parseTimeToMinutes(value) {
    const text = String(value || '').trim();
    const match = text.match(/^(\d{2}):(\d{2})$/);
    if (!match) return null;
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
    return (hours * 60) + minutes;
}

function isScheduledTimeDue(now, sendTime) {
    const scheduledMinutes = parseTimeToMinutes(sendTime);
    const currentMinutes = parseTimeToMinutes(getLagosTimeKey(now));
    if (scheduledMinutes == null || currentMinutes == null) return false;
    return currentMinutes >= scheduledMinutes;
}

function resolveAutoScheduledScope(now = new Date()) {
    const weekday = getLagosWeekdayKey(now);
    const previousDateKey = getPreviousDateKeyInLagos(now);

    if (weekday === 'sat' || weekday === 'sun') {
        return { mode: 'skip', reason: 'weekend-no-send' };
    }

    if (weekday === 'mon') {
        return {
            mode: 'date_range',
            rangeStartDateKey: shiftDateKey(previousDateKey, -2),
            rangeEndDateKey: previousDateKey
        };
    }

    return {
        mode: 'single_day',
        reportDateKey: previousDateKey
    };
}

function getRunEventTimestamp(run = {}) {
    return run.completedAt || run.failedAt || run.updatedAt || run.startedAt || null;
}

function summarizeRunType(trigger = '') {
    const text = String(trigger || '').trim().toLowerCase();
    if (text === 'auto' || text === 'auto_weekend_rollup') return 'auto';
    if (text.startsWith('manual:') || text.startsWith('manual_custom')) return 'manual';
    return text ? 'other' : 'unknown';
}

function serializeRunSummary(run = {}) {
    const eventTimestamp = getRunEventTimestamp(run);
    return {
        runKey: String(run.runKey || '').trim(),
        reportDateKey: String(run.reportDateKey || '').trim(),
        rangeStartDateKey: String(run.rangeStartDateKey || '').trim(),
        rangeEndDateKey: String(run.rangeEndDateKey || '').trim(),
        reportLabel: String(run.reportLabel || '').trim(),
        status: String(run.status || '').trim(),
        trigger: String(run.trigger || '').trim(),
        triggerType: summarizeRunType(run.trigger),
        resendRequested: run.resendRequested === true,
        resendCount: Number(run.resendCount || 0),
        sendTime: String(run.sendTime || '').trim(),
        subject: String(run.subject || '').trim(),
        attachmentFileName: String(run.attachmentFileName || '').trim(),
        error: String(run.error || '').trim(),
        sentCount: Number(run.sentCount || 0),
        failedCount: Number(run.failedCount || 0),
        recipientsCount: Array.isArray(run.recipients) ? run.recipients.length : 0,
        eventDateKey: getLagosDateKeyFromValue(eventTimestamp),
        eventTime: formatLagosDateTime(eventTimestamp)
    };
}

async function loadScheduledReportStatusSnapshot() {
    const settingsSnap = await adminDb.collection('settings').doc('system').get();
    const settings = settingsSnap.exists ? (settingsSnap.data() || {}) : {};
    const scheduled = settings.scheduledReportEmail || {};
    const reportDateKey = getPreviousDateKeyInLagos();
    const todayDateKey = getLagosDateKey();
    const [runSnap, recentRunsSnap] = await Promise.all([
        adminDb.collection('scheduledReportRuns').doc(reportDateKey).get(),
        adminDb.collection('scheduledReportRuns').orderBy('updatedAt', 'desc').limit(20).get()
    ]);

    const lastRun = runSnap.exists ? (runSnap.data() || {}) : null;
    const recentRuns = recentRunsSnap.docs
        .map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }))
        .map((run) => serializeRunSummary(run));
    const todayRuns = recentRuns.filter((run) => run.eventDateKey === todayDateKey);
    const sentStatuses = new Set(['sent', 'partial']);

    return {
        ok: true,
        reportDateKey,
        currentLagosDateKey: todayDateKey,
        currentLagosTime: getLagosTimeKey(),
        enabled: scheduled.enabled === true,
        emailJsConfigured: isEmailJsConfigured(),
        sendTime: String(scheduled.sendTime || '08:00').trim() || '08:00',
        reportDateMode: String(scheduled.reportDateMode || 'previous_day').trim() || 'previous_day',
        recipients: normalizeEmailList(scheduled.recipients || []),
        lastRun: lastRun ? serializeRunSummary(lastRun) : null,
        todaySummary: {
            hasSentForReportDate: sentStatuses.has(String(lastRun?.status || '').toLowerCase()),
            hasAnyRunToday: todayRuns.length > 0,
            totalRuns: todayRuns.length,
            autoRuns: todayRuns.filter((run) => run.triggerType === 'auto').length,
            manualRuns: todayRuns.filter((run) => run.triggerType === 'manual').length,
            resendRuns: todayRuns.filter((run) => run.resendRequested === true).length,
            successfulRuns: todayRuns.filter((run) => sentStatuses.has(String(run.status || '').toLowerCase())).length,
            failedRuns: todayRuns.filter((run) => String(run.status || '').toLowerCase() === 'failed').length
        },
        todayRuns,
        recentRuns
    };
}

async function resolveUserMapByEmail(emails) {
    const out = new Map();
    for (const email of emails) {
        const normalized = safeLowerEmail(email);
        if (!normalized) continue;
        const snap = await adminDb.collection('users').where('email', '==', normalized).limit(1).get();
        if (!snap.empty) out.set(normalized, { id: snap.docs[0].id, ...(snap.docs[0].data() || {}) });
    }
    if (out.size < emails.length) {
        const targets = new Set(emails.map((e) => safeLowerEmail(e)).filter(Boolean));
        const allSnap = await adminDb.collection('users').get();
        allSnap.docs.forEach((docSnap) => {
            const data = docSnap.data() || {};
            const normalized = safeLowerEmail(data.email);
            if (!normalized || !targets.has(normalized)) return;
            if (!out.has(normalized)) out.set(normalized, { id: docSnap.id, ...data });
        });
    }
    return out;
}

function corsOptionsDelegate(req, callback) {
    if (allowedOrigins.length === 0) {
        callback(null, { origin: true });
        return;
    }

    const reqOrigin = req.header('Origin');
    const allowed = reqOrigin && allowedOrigins.includes(reqOrigin);
    callback(null, { origin: Boolean(allowed) });
}

async function authMiddleware(req, res, next) {
    if (!requireAuth) return next();

    const authHeader = String(req.header('Authorization') || '');
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    if (!token) {
        return res.status(401).json({ ok: false, error: 'Missing Firebase ID token' });
    }

    try {
        const decoded = await admin.auth().verifyIdToken(token);
        req.user = decoded;
        return next();
    } catch (err) {
        return res.status(401).json({ ok: false, error: 'Invalid Firebase ID token' });
    }
}

initFirebaseAdmin();
const adminDb = admin.firestore();

const sendgridApiKey = String(process.env.SENDGRID_API_KEY || '').trim();
if (sendgridApiKey) {
    sgMail.setApiKey(sendgridApiKey);
}

const defaultFromEmail = normalizeEmail(process.env.SENDGRID_FROM_EMAIL);

app.use(cors(corsOptionsDelegate));
app.use(express.json({ limit: '256kb' }));

app.get('/health', (_req, res) => {
    res.json({
        ok: true,
        service: 'rsa-render-email-api',
        authRequired: requireAuth,
        ts: new Date().toISOString()
    });
});

app.get('/api/server-time', (_req, res) => {
    const now = new Date();
    res.json({
        ok: true,
        epochMs: now.getTime(),
        iso: now.toISOString(),
        timeZone: 'Africa/Lagos',
        lagosDateTime: lagosDateTimeFormatter.format(now),
        lagosDateKey: lagosDateKeyFormatter.format(now)
    });
});

function isEmailJsConfigured() {
    return Boolean(EMAILJS_PUBLIC_KEY && EMAILJS_SERVICE_ID && EMAILJS_REPORT_TEMPLATE_ID);
}

async function sendEmailViaEmailJs({ to, subject, text, reportDateKey, reportIntro, attachmentDataUri, attachmentFileName }) {
    if (!isEmailJsConfigured()) {
        throw new Error('EmailJS is not configured on this service');
    }

    const templateParams = {
        to_email: to,
        email_subject: subject,
        email_message: text,
        reply_to: 'no-reply@cmbankrsa.com',
        report_date: reportDateKey,
        report_title: 'Daily Report',
        report_intro: String(reportIntro || 'Your operational report is ready and attached for review.').trim(),
        attachment_note: 'The Excel report workbook is attached to this email for download and review.',
        portal_name: 'CMBank RSA Portal',
        [EMAILJS_ATTACHMENT_PARAM]: attachmentDataUri,
        [EMAILJS_ATTACHMENT_FILENAME_PARAM]: attachmentFileName
    };

    const payload = {
        service_id: EMAILJS_SERVICE_ID,
        template_id: EMAILJS_REPORT_TEMPLATE_ID,
        user_id: EMAILJS_PUBLIC_KEY,
        template_params: templateParams
    };

    if (EMAILJS_PRIVATE_KEY) {
        payload.accessToken = EMAILJS_PRIVATE_KEY;
    }

    const response = await fetch('https://api.emailjs.com/api/v1.0/email/send', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        const body = await response.text().catch(() => '');
        throw new Error(body || `EmailJS returned ${response.status}`);
    }
}

function normalizeDateKey(value) {
    const text = String(value || '').trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : '';
}

function resolveManualReportScope({ reportDateKey, rangeStartDateKey, rangeEndDateKey } = {}) {
    const singleDate = normalizeDateKey(reportDateKey);
    const rangeStart = normalizeDateKey(rangeStartDateKey);
    const rangeEnd = normalizeDateKey(rangeEndDateKey);
    if (rangeStart && rangeEnd) {
        const start = rangeStart <= rangeEnd ? rangeStart : rangeEnd;
        const end = rangeStart <= rangeEnd ? rangeEnd : rangeStart;
        return {
            mode: 'date_range',
            reportDateKey: '',
            rangeStartDateKey: start,
            rangeEndDateKey: end,
            label: `${start} to ${end}`,
            runKey: `range_${start}_to_${end}`
        };
    }
    return {
        mode: 'single_day',
        reportDateKey: singleDate || getPreviousDateKeyInLagos(),
        rangeStartDateKey: '',
        rangeEndDateKey: '',
        label: singleDate || getPreviousDateKeyInLagos(),
        runKey: singleDate || getPreviousDateKeyInLagos()
    };
}

async function reserveScheduledReportRun(runKey, trigger, forceResend = false) {
    const runRef = adminDb.collection('scheduledReportRuns').doc(runKey);
    let shouldSend = false;
    await adminDb.runTransaction(async (tx) => {
        const snap = await tx.get(runRef);
        const data = snap.exists ? (snap.data() || {}) : {};
        if (!forceResend && String(data.status || '').toLowerCase() === 'sent') {
            return;
        }
        shouldSend = true;
        tx.set(runRef, {
            runKey,
            status: 'sending',
            trigger,
            resendRequested: forceResend === true,
            updatedAt: admin.firestore.FieldValue.serverTimestamp(),
            startedAt: data.startedAt || admin.firestore.FieldValue.serverTimestamp(),
            resentAt: forceResend === true ? admin.firestore.FieldValue.serverTimestamp() : data.resentAt || null,
            resendCount: forceResend === true ? Number(data.resendCount || 0) + 1 : Number(data.resendCount || 0)
        }, { merge: true });
    });
    return { shouldSend, runRef };
}

async function sendScheduledReportForDate({ reportDateKey, trigger = 'manual', forceResend = false }) {
    const scope = resolveManualReportScope({ reportDateKey });
    const normalizedDateKey = scope.reportDateKey || getPreviousDateKeyInLagos();
    const settingsSnap = await adminDb.collection('settings').doc('system').get();
    const settings = settingsSnap.exists ? (settingsSnap.data() || {}) : {};
    const scheduled = settings.scheduledReportEmail || {};
    const enabled = scheduled.enabled === true;
    const sendTime = String(scheduled.sendTime || '08:00').trim() || '08:00';
    const subject = String(scheduled.subject || 'Daily RSA Report').trim() || 'Daily RSA Report';
    const bodyText = String(scheduled.body || 'Hello,\n\nPlease find the attached daily RSA report.\n\nRegards,\nCMBank RSA Portal').trim();
    const recipients = normalizeEmailList(scheduled.recipients || []);
    const reportDateMode = String(scheduled.reportDateMode || 'previous_day').trim() || 'previous_day';

    if (!enabled) {
        return { ok: false, skipped: true, reason: 'scheduled-report-disabled' };
    }
    if (reportDateMode !== 'previous_day') {
        return { ok: false, skipped: true, reason: 'unsupported-report-date-mode' };
    }
    if (!recipients.length) {
        return { ok: false, skipped: true, reason: 'no-recipients' };
    }

    const { shouldSend, runRef } = await reserveScheduledReportRun(scope.runKey, trigger, forceResend);
    if (!shouldSend) {
        return { ok: true, skipped: true, reason: 'already-sent', reportDateKey: normalizedDateKey, reportLabel: scope.label };
    }

    try {
        const [usersSnap, submissionsSnap] = await Promise.all([
            adminDb.collection('users').get(),
            adminDb.collection('submissions').get()
        ]);
        const users = usersSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        const submissions = submissionsSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        const workbookBuffer = Buffer.from(await createDailyReportWorkbookBuffer({
            submissions,
            users,
            reportDateKey: normalizedDateKey
        }));
        const attachmentFileName = `cmbank_daily_report_${normalizedDateKey}.xlsx`;
        const attachmentDataUri = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${workbookBuffer.toString('base64')}`;
        const text = bodyText;

        let sentCount = 0;
        const failures = [];
        for (const recipient of recipients) {
            try {
                await sendEmailViaEmailJs({
                    to: recipient,
                    subject: `${subject} - ${normalizedDateKey}`,
                    text,
                    reportDateKey: normalizedDateKey,
                    reportIntro: 'Your previous-day operational report is ready and attached for review.',
                    attachmentDataUri,
                    attachmentFileName
                });
                sentCount += 1;
            } catch (err) {
                failures.push({ recipient, error: String(err?.message || 'send-failed') });
            }
        }

        const finalStatus = sentCount > 0
            ? (failures.length ? 'partial' : 'sent')
            : 'failed';
        await runRef.set({
            status: finalStatus,
            runKey: scope.runKey,
            reportDateKey: normalizedDateKey,
            reportLabel: scope.label,
            trigger,
            resendRequested: forceResend === true,
            sendTime,
            recipients,
            sentCount,
            failedCount: failures.length,
            failures,
            attachmentFileName,
            completedAt: admin.firestore.FieldValue.serverTimestamp(),
            subject,
            error: admin.firestore.FieldValue.delete(),
            failedAt: admin.firestore.FieldValue.delete()
        }, { merge: true });

        return {
            ok: sentCount > 0 && failures.length === 0,
            reportDateKey: normalizedDateKey,
            reportLabel: scope.label,
            attachmentFileName,
            sentCount,
            failedCount: failures.length,
            failures
        };
    } catch (err) {
        await runRef.set({
            status: 'failed',
            runKey: scope.runKey,
            reportDateKey: normalizedDateKey,
            trigger,
            error: String(err?.message || 'scheduled-report-failed'),
            failedAt: admin.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
        throw err;
    }
}

async function sendManualReport({ reportDateKey, rangeStartDateKey, rangeEndDateKey, trigger = 'manual_custom', forceResend = false }) {
    const scope = resolveManualReportScope({ reportDateKey, rangeStartDateKey, rangeEndDateKey });
    const settingsSnap = await adminDb.collection('settings').doc('system').get();
    const settings = settingsSnap.exists ? (settingsSnap.data() || {}) : {};
    const scheduled = settings.scheduledReportEmail || {};
    const enabled = scheduled.enabled === true;
    const sendTime = String(scheduled.sendTime || '08:00').trim() || '08:00';
    const subject = String(scheduled.subject || 'Daily RSA Report').trim() || 'Daily RSA Report';
    const bodyText = String(scheduled.body || 'Hello,\n\nPlease find the attached daily RSA report.\n\nRegards,\nCMBank RSA Portal').trim();
    const recipients = normalizeEmailList(scheduled.recipients || []);

    if (!enabled) {
        return { ok: false, skipped: true, reason: 'scheduled-report-disabled' };
    }
    if (!recipients.length) {
        return { ok: false, skipped: true, reason: 'no-recipients' };
    }

    const { shouldSend, runRef } = await reserveScheduledReportRun(scope.runKey, trigger, forceResend);
    if (!shouldSend) {
        return { ok: true, skipped: true, reason: 'already-sent', reportDateKey: scope.reportDateKey, reportLabel: scope.label };
    }

    try {
        const [usersSnap, submissionsSnap] = await Promise.all([
            adminDb.collection('users').get(),
            adminDb.collection('submissions').get()
        ]);
        const users = usersSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        const submissions = submissionsSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        const workbookBuffer = Buffer.from(await createDailyReportWorkbookBuffer({
            submissions,
            users,
            reportDateKey: scope.reportDateKey,
            rangeStartDateKey: scope.rangeStartDateKey,
            rangeEndDateKey: scope.rangeEndDateKey
        }));
        const attachmentSuffix = scope.mode === 'date_range'
            ? `${scope.rangeStartDateKey}_to_${scope.rangeEndDateKey}`
            : scope.reportDateKey;
        const attachmentFileName = `cmbank_daily_report_${attachmentSuffix}.xlsx`;
        const attachmentDataUri = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${workbookBuffer.toString('base64')}`;
        const text = bodyText;

        let sentCount = 0;
        const failures = [];
        for (const recipient of recipients) {
            try {
                await sendEmailViaEmailJs({
                    to: recipient,
                    subject: `${subject} - ${scope.label}`,
                    text,
                    reportDateKey: scope.label,
                    reportIntro: scope.mode === 'date_range'
                        ? 'Your selected report range is ready and attached for review.'
                        : 'Your selected report day is ready and attached for review.',
                    attachmentDataUri,
                    attachmentFileName
                });
                sentCount += 1;
            } catch (err) {
                failures.push({ recipient, error: String(err?.message || 'send-failed') });
            }
        }

        const finalStatus = sentCount > 0
            ? (failures.length ? 'partial' : 'sent')
            : 'failed';
        await runRef.set({
            status: finalStatus,
            runKey: scope.runKey,
            reportDateKey: scope.reportDateKey,
            rangeStartDateKey: scope.rangeStartDateKey,
            rangeEndDateKey: scope.rangeEndDateKey,
            reportLabel: scope.label,
            trigger,
            resendRequested: forceResend === true,
            sendTime,
            recipients,
            sentCount,
            failedCount: failures.length,
            failures,
            attachmentFileName,
            completedAt: admin.firestore.FieldValue.serverTimestamp(),
            subject,
            error: admin.firestore.FieldValue.delete(),
            failedAt: admin.firestore.FieldValue.delete()
        }, { merge: true });

        return {
            ok: sentCount > 0 && failures.length === 0,
            mode: scope.mode,
            reportDateKey: scope.reportDateKey,
            rangeStartDateKey: scope.rangeStartDateKey,
            rangeEndDateKey: scope.rangeEndDateKey,
            reportLabel: scope.label,
            attachmentFileName,
            sentCount,
            failedCount: failures.length,
            failures
        };
    } catch (err) {
        await runRef.set({
            status: 'failed',
            runKey: scope.runKey,
            reportDateKey: scope.reportDateKey,
            rangeStartDateKey: scope.rangeStartDateKey,
            rangeEndDateKey: scope.rangeEndDateKey,
            reportLabel: scope.label,
            trigger,
            error: String(err?.message || 'scheduled-report-failed'),
            failedAt: admin.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
        throw err;
    }
}

async function requireAdminRole(req, res, next) {
    if (!req.user?.uid && !req.user?.email) {
        return res.status(401).json({ ok: false, error: 'Missing authenticated user context' });
    }

    try {
        let role = '';
        let status = '';

        if (req.user?.uid) {
            const byUid = await adminDb.collection('users').where('uid', '==', req.user.uid).limit(1).get();
            if (!byUid.empty) {
                const data = byUid.docs[0].data() || {};
                role = normalizeRole(data.role);
                status = String(data.status || '').toLowerCase();
            }
        }

        if (!role && req.user?.email) {
            const normalizedEmail = normalizeEmail(req.user.email);
            const byEmail = await adminDb.collection('users').where('email', '==', normalizedEmail).limit(1).get();
            if (!byEmail.empty) {
                const data = byEmail.docs[0].data() || {};
                role = normalizeRole(data.role);
                status = String(data.status || '').toLowerCase();
            }
        }

        if (role !== 'admin' || (status && status !== 'active')) {
            return res.status(403).json({ ok: false, error: 'Admin access required' });
        }

        return next();
    } catch (err) {
        return res.status(500).json({ ok: false, error: 'Failed to validate admin role' });
    }
}

async function requireAdminOrSuperAdminRole(req, res, next) {
    if (!req.user?.uid && !req.user?.email) {
        return res.status(401).json({ ok: false, error: 'Missing authenticated user context' });
    }

    try {
        let role = '';
        let status = '';

        if (req.user?.uid) {
            const byUid = await adminDb.collection('users').where('uid', '==', req.user.uid).limit(1).get();
            if (!byUid.empty) {
                const data = byUid.docs[0].data() || {};
                role = normalizeRole(data.role);
                status = String(data.status || '').toLowerCase();
            }
        }

        if (!role && req.user?.email) {
            const normalizedEmail = normalizeEmail(req.user.email);
            const byEmail = await adminDb.collection('users').where('email', '==', normalizedEmail).limit(1).get();
            if (!byEmail.empty) {
                const data = byEmail.docs[0].data() || {};
                role = normalizeRole(data.role);
                status = String(data.status || '').toLowerCase();
            }
        }

        if (!['admin', 'super_admin'].includes(role) || (status && status !== 'active')) {
            return res.status(403).json({ ok: false, error: 'Admin access required' });
        }

        return next();
    } catch (err) {
        return res.status(500).json({ ok: false, error: 'Failed to validate admin role' });
    }
}

app.post('/api/send-alert', authMiddleware, async (req, res) => {
    try {
        if (!sendgridApiKey || !defaultFromEmail) {
            return res.status(503).json({
                ok: false,
                error: 'SendGrid is not configured on this service'
            });
        }

        const eventKey = String(req.body?.eventKey || '').trim();
        const to = normalizeEmail(req.body?.to);
        const subject = String(req.body?.subject || '').trim();
        const text = String(req.body?.text || '').trim();
        const html = String(req.body?.html || '').trim();
        const meta = req.body?.meta || {};
        const type = String(meta?.type || '').trim();

        if (!eventKey || !to || !subject) {
            return res.status(400).json({ ok: false, error: 'eventKey, to, and subject are required' });
        }
        if (!allowedTypes.has(type)) {
            return res.status(400).json({ ok: false, error: 'Unsupported email type' });
        }

        const [sendResult] = await sgMail.send({
            to,
            from: defaultFromEmail,
            subject,
            text: text || undefined,
            html: html || undefined,
            customArgs: {
                eventKey,
                type
            }
        });

        return res.json({
            ok: true,
            messageId: String(sendResult?.headers?.['x-message-id'] || '')
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send alert email')
        });
    }
});

app.get('/api/scheduled-report/status', authMiddleware, requireAdminOrSuperAdminRole, async (_req, res) => {
    try {
        return res.json(await loadScheduledReportStatusSnapshot());
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to load scheduled report status')
        });
    }
});

app.post('/api/scheduled-report/send-now', authMiddleware, requireAdminOrSuperAdminRole, async (req, res) => {
    try {
        const forceResend = req.body?.forceResend === true;
        const reportDateKey = normalizeDateKey(req.body?.reportDateKey);
        const rangeStartDateKey = normalizeDateKey(req.body?.rangeStartDateKey);
        const rangeEndDateKey = normalizeDateKey(req.body?.rangeEndDateKey);
        const result = (reportDateKey || (rangeStartDateKey && rangeEndDateKey))
            ? await sendManualReport({
                reportDateKey,
                rangeStartDateKey,
                rangeEndDateKey,
                trigger: `manual:${normalizeEmail(req.user?.email) || 'admin'}`,
                forceResend
            })
            : await sendScheduledReportForDate({
                reportDateKey: getPreviousDateKeyInLagos(),
                trigger: `manual:${normalizeEmail(req.user?.email) || 'admin'}`,
                forceResend
            });
        if (result?.skipped) {
            const reason = String(result?.reason || 'scheduled-report-skipped');
            const statusCode = reason === 'already-sent' ? 409 : 400;
            return res.status(statusCode).json({ ok: false, ...result, error: reason });
        }
        return res.json({ ok: true, ...result });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send scheduled report')
        });
    }
});

app.post('/api/public/password-reset-request', async (req, res) => {
    try {
        if (!sendgridApiKey || !defaultFromEmail) {
            return res.status(503).json({
                ok: false,
                error: 'SendGrid is not configured on this service'
            });
        }

        const email = normalizeEmail(req.body?.email);
        const resetPageUrl = buildResetPageUrl(req, req.body?.resetPageUrl);

        if (!email) {
            return res.status(400).json({ ok: false, error: 'Valid email is required' });
        }
        if (!resetPageUrl) {
            return res.status(400).json({ ok: false, error: 'Reset page URL is not configured' });
        }

        let userRecord = null;
        try {
            userRecord = await admin.auth().getUserByEmail(email);
        } catch (err) {
            if (String(err?.code || '') !== 'auth/user-not-found') {
                throw err;
            }
        }

        if (!userRecord) {
            return res.json({
                ok: true,
                message: 'If the account exists, a reset link will be sent.'
            });
        }

        const generatedLink = await admin.auth().generatePasswordResetLink(email);
        const generatedUrl = new URL(generatedLink);
        const oobCode = String(generatedUrl.searchParams.get('oobCode') || '').trim();
        const mode = String(generatedUrl.searchParams.get('mode') || 'resetPassword').trim();
        const apiKey = String(generatedUrl.searchParams.get('apiKey') || '').trim();
        const lang = String(generatedUrl.searchParams.get('lang') || 'en').trim();

        if (!oobCode) {
            throw new Error('Generated password reset link is missing oobCode');
        }

        const customResetUrl = new URL(resetPageUrl);
        customResetUrl.searchParams.set('mode', mode);
        customResetUrl.searchParams.set('oobCode', oobCode);
        if (apiKey) customResetUrl.searchParams.set('apiKey', apiKey);
        if (lang) customResetUrl.searchParams.set('lang', lang);

        const safeEmail = escapeHtml(email);
        const safeResetUrl = escapeHtml(customResetUrl.toString());

        await sgMail.send({
            to: email,
            from: defaultFromEmail,
            subject: 'Reset your CMBank RSA password',
            text: [
                'We received a request to reset your CMBank RSA password.',
                '',
                `Open this link to continue: ${customResetUrl.toString()}`,
                '',
                'If you did not request this, you can ignore this email.'
            ].join('\n'),
            html: `
                <div style="font-family:Segoe UI,Arial,sans-serif;background:#f5f8fc;padding:24px;">
                    <div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #dbe5f0;border-radius:18px;overflow:hidden;box-shadow:0 16px 40px rgba(0,51,102,0.12);">
                        <div style="background:linear-gradient(135deg,#003366,#0b5cab);padding:24px;color:#ffffff;">
                            <div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;font-weight:700;opacity:.9;">CMBank RSA Portal</div>
                            <h1 style="margin:10px 0 0;font-size:24px;line-height:1.25;">Reset your password</h1>
                            <p style="margin:10px 0 0;font-size:14px;line-height:1.6;color:#dbeafe;">A password reset was requested for ${safeEmail}.</p>
                        </div>
                        <div style="padding:24px;color:#0f172a;">
                            <p style="margin:0 0 16px;font-size:14px;line-height:1.7;">Use the button below to open the CMBank RSA reset page and choose a new password.</p>
                            <p style="margin:0 0 22px;">
                                <a href="${safeResetUrl}" style="display:inline-block;background:linear-gradient(135deg,#003366,#0b5cab);color:#ffffff;text-decoration:none;padding:12px 18px;border-radius:10px;font-size:14px;font-weight:700;">Open Reset Page</a>
                            </p>
                            <p style="margin:0 0 12px;font-size:13px;line-height:1.7;color:#475569;">If the button does not open, copy and paste this link into your browser:</p>
                            <p style="margin:0;font-size:12px;line-height:1.7;word-break:break-all;color:#0b5cab;">${safeResetUrl}</p>
                        </div>
                        <div style="padding:18px 24px;background:#f8fbff;border-top:1px solid #e5edf6;color:#64748b;font-size:12px;line-height:1.6;">
                            If you did not request this reset, you can safely ignore this email.
                        </div>
                    </div>
                </div>
            `
        });

        return res.json({
            ok: true,
            message: 'If the account exists, a reset link will be sent.'
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send password reset email')
        });
    }
});

app.post('/api/admin/reset-password', authMiddleware, requireAdminRole, async (req, res) => {
    try {
        const userId = String(req.body?.userId || '').trim();
        const newPassword = String(req.body?.newPassword || '123456').trim();

        if (!userId) {
            return res.status(400).json({ ok: false, error: 'userId is required' });
        }
        if (newPassword.length < 6) {
            return res.status(400).json({ ok: false, error: 'Password must be at least 6 characters' });
        }

        const userRef = adminDb.collection('users').doc(userId);
        const userSnap = await userRef.get();
        if (!userSnap.exists) {
            return res.status(404).json({ ok: false, error: 'Target user not found' });
        }

        const userData = userSnap.data() || {};
        const targetUid = String(userData.uid || '').trim();
        const targetEmail = normalizeEmail(userData.email);

        if (!targetUid && !targetEmail) {
            return res.status(400).json({ ok: false, error: 'Target user has no uid/email linked to Auth' });
        }

        let authUid = targetUid;
        if (!authUid && targetEmail) {
            const userRecord = await admin.auth().getUserByEmail(targetEmail);
            authUid = userRecord.uid;
        }

        await admin.auth().updateUser(authUid, { password: newPassword });

        await userRef.set({
            passwordResetAt: admin.firestore.FieldValue.serverTimestamp(),
            passwordResetBy: req.user?.email || req.user?.uid || 'admin-api',
            passwordResetMethod: 'admin_direct',
            passwordResetDefault: true
        }, { merge: true });

        await adminDb.collection('audit').add({
            action: 'user_password_reset_direct',
            userId,
            userEmail: targetEmail || '',
            resetTo: '123456',
            performedBy: req.user?.email || req.user?.uid || 'admin-api',
            timestamp: admin.firestore.FieldValue.serverTimestamp()
        });

        return res.json({
            ok: true,
            message: 'Password reset to default (123456)'
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to reset password')
        });
    }
});

app.post('/api/push/register-token', authMiddleware, async (req, res) => {
    try {
        const token = String(req.body?.token || '').trim();
        const previousToken = String(req.body?.previousToken || '').trim();
        const profileDocId = String(req.body?.profileDocId || '').trim();
        const senderEmail = safeLowerEmail(req.user?.email);
        const senderUid = String(req.user?.uid || '').trim();

        if (!token) {
            return res.status(400).json({ ok: false, error: 'token is required' });
        }
        if (!senderEmail && !senderUid) {
            return res.status(401).json({ ok: false, error: 'Authenticated user is required' });
        }

        let currentUserRef = null;
        if (profileDocId) {
            const candidateRef = adminDb.collection('users').doc(profileDocId);
            const candidateSnap = await candidateRef.get();
            if (candidateSnap.exists) {
                const data = candidateSnap.data() || {};
                const candidateEmail = safeLowerEmail(data.email);
                const candidateUid = String(data.uid || '').trim();
                if ((senderEmail && candidateEmail === senderEmail) || (senderUid && candidateUid === senderUid)) {
                    currentUserRef = candidateRef;
                }
            }
        }

        if (!currentUserRef && senderEmail) {
            const snap = await adminDb.collection('users').where('email', '==', senderEmail).limit(1).get();
            if (!snap.empty) currentUserRef = snap.docs[0].ref;
        }

        if (!currentUserRef && senderUid) {
            const snap = await adminDb.collection('users').where('uid', '==', senderUid).limit(1).get();
            if (!snap.empty) currentUserRef = snap.docs[0].ref;
        }

        if (!currentUserRef) {
            return res.status(404).json({ ok: false, error: 'User profile not found' });
        }

        const tokensToPrune = [];
        if (token) tokensToPrune.push(token);
        if (previousToken && previousToken !== token) tokensToPrune.push(previousToken);

        for (const tokenToPrune of Array.from(new Set(tokensToPrune))) {
            const ownersSnap = await adminDb.collection('users').where('fcmTokens', 'array-contains', tokenToPrune).get();
            await Promise.all(ownersSnap.docs.map((docSnap) => (
                docSnap.ref.update({
                    fcmTokens: admin.firestore.FieldValue.arrayRemove(tokenToPrune),
                    fcmTokenPrunedAt: admin.firestore.FieldValue.serverTimestamp()
                }).catch(() => {})
            )));
        }

        await currentUserRef.set({
            fcmTokens: admin.firestore.FieldValue.arrayUnion(token),
            fcmLastTokenAt: admin.firestore.FieldValue.serverTimestamp(),
            fcmLastTokenPlatform: String(req.body?.platform || '').trim()
        }, { merge: true });

        return res.json({ ok: true });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to register push token')
        });
    }
});

app.post('/api/chat/push', authMiddleware, async (req, res) => {
    try {
        const submissionId = String(req.body?.submissionId || '').trim();
        const customerName = String(req.body?.customerName || '').trim();
        const messageText = String(req.body?.messageText || '').trim();
        if (!submissionId || !messageText) {
            return res.status(400).json({ ok: false, error: 'submissionId and messageText are required' });
        }

        const senderEmail = safeLowerEmail(req.user?.email);
        if (!senderEmail) {
            return res.status(401).json({ ok: false, error: 'Authenticated email is required' });
        }

        const subRef = adminDb.collection('submissions').doc(submissionId);
        const subSnap = await subRef.get();
        if (!subSnap.exists) {
            return res.status(404).json({ ok: false, error: 'Submission not found' });
        }
        const sub = subSnap.data() || {};
        const chatSnap = await adminDb.collection('applicationChats').doc(submissionId).get();
        const chatMeta = chatSnap.exists ? (chatSnap.data() || {}) : {};

        // Strict live participant set only from current submission workflow.
        // Avoid stale historical chat/admin participant lists.
        const participantEmails = getSubmissionParticipantEmails(sub);

        const senderCanAccess = participantEmails.includes(senderEmail);
        if (!senderCanAccess) {
            const senderUserSnap = await adminDb.collection('users').where('email', '==', senderEmail).limit(1).get();
            const senderRole = !senderUserSnap.empty ? normalizeRole(senderUserSnap.docs[0].data()?.role) : '';
            if (!['admin', 'super_admin'].includes(senderRole)) {
                return res.status(403).json({ ok: false, error: 'Not permitted for this submission' });
            }
        }

        const recipientEmailsRaw = getChatPushRecipientEmails(sub, chatMeta).filter((e) => e !== senderEmail);
        let recipientEmails = [...recipientEmailsRaw];
        if (recipientEmails.length === 0) {
            return res.json({
                ok: true,
                sent: 0,
                reason: 'no-recipients',
                debug: {
                    senderEmail,
                    participantEmails,
                    recipientEmails
                }
            });
        }

        const usersByEmail = new Map();
        for (const email of recipientEmails) {
            const uSnap = await adminDb.collection('users').where('email', '==', email).limit(1).get();
            if (!uSnap.empty) usersByEmail.set(email, { id: uSnap.docs[0].id, ...(uSnap.docs[0].data() || {}) });
        }

        // Fallback for legacy/unclean email casing or spacing in user docs.
        if (usersByEmail.size < recipientEmails.length) {
            const allUsersSnap = await adminDb.collection('users').get();
            const normalizedTarget = new Set(recipientEmails.map((e) => safeLowerEmail(e)).filter(Boolean));
            allUsersSnap.docs.forEach((docSnap) => {
                const data = docSnap.data() || {};
                const normalized = safeLowerEmail(data.email);
                if (!normalized || !normalizedTarget.has(normalized)) return;
                if (!usersByEmail.has(normalized)) {
                    usersByEmail.set(normalized, { id: docSnap.id, ...data });
                }
            });
        }

        // Exclude super admin and inactive accounts from chat notifications/recipients.
        recipientEmails = recipientEmails.filter((email) => {
            const user = usersByEmail.get(email);
            const role = normalizeRole(user?.role);
            const status = String(user?.status || '').toLowerCase();
            return role !== 'super_admin' && status !== 'deactivated' && status !== 'pending';
        });

        if (recipientEmails.length === 0) {
            return res.json({
                ok: true,
                sent: 0,
                reason: 'no-recipients-after-role-filter',
                debug: {
                    senderEmail,
                    participantEmails,
                    recipientEmailsRaw,
                    recipientEmails,
                    usersMatched: Array.from(usersByEmail.keys())
                }
            });
        }

        const tokenOwners = [];
        for (const email of recipientEmails) {
            const userData = usersByEmail.get(email);
            const tokens = Array.isArray(userData?.fcmTokens) ? userData.fcmTokens : [];
            for (const token of tokens) {
                const t = String(token || '').trim();
                if (t) tokenOwners.push({ email, userId: userData?.id || '', token: t });
            }
        }

        const uniqueMap = new Map();
        tokenOwners.forEach((entry) => { if (!uniqueMap.has(entry.token)) uniqueMap.set(entry.token, entry); });
        const uniqueOwners = Array.from(uniqueMap.values());
        const tokens = uniqueOwners.map((x) => x.token);
        if (tokens.length === 0) {
            return res.json({
                ok: true,
                sent: 0,
                reason: 'no-device-tokens',
                debug: {
                    senderEmail,
                    participantEmails,
                    recipientEmailsRaw,
                    recipientEmails,
                    usersMatched: Array.from(usersByEmail.keys()),
                    tokenOwners: uniqueOwners.map((x) => x.email)
                }
            });
        }

        const title = `New Chat Message - ${customerName || 'Application'}`;
        const body = messageText.slice(0, 120);
        const clickUrl = `/dashboard.html?chat=${encodeURIComponent(submissionId)}`;

        const response = await admin.messaging().sendEachForMulticast({
            tokens,
            data: {
                submissionId,
                clickUrl,
                title,
                body
            },
            webpush: {
                fcmOptions: { link: clickUrl }
            }
        });

        const badTokens = [];
        const failedCodes = [];
        response.responses.forEach((r, idx) => {
            if (r.success) return;
            const code = String(r.error?.code || '');
            if (code) failedCodes.push(code);
            if (
                code.includes('registration-token-not-registered') ||
                code.includes('invalid-argument')
            ) {
                badTokens.push(uniqueOwners[idx]);
            }
        });

        for (const bad of badTokens) {
            if (!bad?.userId) continue;
            try {
                const ref = adminDb.collection('users').doc(bad.userId);
                const snap = await ref.get();
                if (!snap.exists) continue;
                const arr = Array.isArray(snap.data()?.fcmTokens) ? snap.data().fcmTokens : [];
                const next = arr.filter((t) => String(t || '').trim() !== bad.token);
                await ref.set({ fcmTokens: next, fcmTokenPrunedAt: admin.firestore.FieldValue.serverTimestamp() }, { merge: true });
            } catch (_) {}
        }

        return res.json({
            ok: true,
            sent: response.successCount,
            failed: response.failureCount,
            debug: {
                senderEmail,
                participantEmails,
                recipientEmailsRaw,
                recipientEmails,
                usersMatched: Array.from(usersByEmail.keys()),
                tokenOwners: uniqueOwners.map((x) => x.email),
                failedCodes
            }
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send chat push')
        });
    }
});

app.post('/api/submission/status-push', authMiddleware, async (req, res) => {
    try {
        const submissionId = String(req.body?.submissionId || '').trim();
        const customerName = String(req.body?.customerName || '').trim();
        const newStatus = String(req.body?.newStatus || '').trim().toLowerCase();
        const statusLabel = String(req.body?.statusLabel || '').trim();
        const actionLabel = String(req.body?.actionLabel || '').trim();
        const customMessage = String(req.body?.message || '').trim();
        if (!submissionId || !newStatus) {
            return res.status(400).json({ ok: false, error: 'submissionId and newStatus are required' });
        }

        const senderEmail = safeLowerEmail(req.user?.email);
        if (!senderEmail) {
            return res.status(401).json({ ok: false, error: 'Authenticated email is required' });
        }

        const subSnap = await adminDb.collection('submissions').doc(submissionId).get();
        if (!subSnap.exists) {
            return res.status(404).json({ ok: false, error: 'Submission not found' });
        }
        const sub = subSnap.data() || {};

        const participantEmails = getSubmissionParticipantEmails(sub);

        const senderCanAccess = participantEmails.includes(senderEmail);
        if (!senderCanAccess) {
            const senderUserSnap = await adminDb.collection('users').where('email', '==', senderEmail).limit(1).get();
            const senderRole = !senderUserSnap.empty ? normalizeRole(senderUserSnap.docs[0].data()?.role) : '';
            if (!['admin'].includes(senderRole)) {
                return res.status(403).json({ ok: false, error: 'Not permitted for this submission' });
            }
        }

        let recipientEmails = getStatusPushRecipientEmails(sub, newStatus).filter((e) => e !== senderEmail);
        if (recipientEmails.length === 0) {
            return res.json({ ok: true, sent: 0, reason: 'no-recipients' });
        }

        const usersByEmail = await resolveUserMapByEmail(recipientEmails);
        recipientEmails = recipientEmails.filter((email) => {
            const u = usersByEmail.get(email);
            const role = normalizeRole(u?.role);
            const status = String(u?.status || '').toLowerCase();
            return role !== 'super_admin' && status !== 'pending' && status !== 'deactivated';
        });
        if (recipientEmails.length === 0) {
            return res.json({ ok: true, sent: 0, reason: 'no-recipients-after-role-filter' });
        }

        const tokenOwners = [];
        for (const email of recipientEmails) {
            const userData = usersByEmail.get(email);
            const tokens = Array.isArray(userData?.fcmTokens) ? userData.fcmTokens : [];
            for (const token of tokens) {
                const t = String(token || '').trim();
                if (t) tokenOwners.push({ email, userId: userData?.id || '', token: t });
            }
        }
        const uniqueMap = new Map();
        tokenOwners.forEach((entry) => { if (!uniqueMap.has(entry.token)) uniqueMap.set(entry.token, entry); });
        const uniqueOwners = Array.from(uniqueMap.values());
        const tokens = uniqueOwners.map((x) => x.token);
        if (tokens.length === 0) {
            return res.json({ ok: true, sent: 0, reason: 'no-device-tokens' });
        }

        const readable = statusLabel || newStatus.replace(/_/g, ' ').replace(/\b\w/g, (m) => m.toUpperCase());
        const actionReadable = actionLabel || 'Status Updated';
        const title = `${actionReadable} - ${customerName || 'Application'}`;
        const body = customMessage || `Application status changed to ${readable}.`;
        const clickUrl = `/dashboard.html?chat=${encodeURIComponent(submissionId)}`;

        const response = await admin.messaging().sendEachForMulticast({
            tokens,
            data: {
                submissionId,
                clickUrl,
                title,
                body,
                newStatus,
                actionLabel: actionReadable
            },
            webpush: {
                fcmOptions: { link: clickUrl }
            }
        });

        return res.json({
            ok: true,
            sent: response.successCount,
            failed: response.failureCount
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send status push')
        });
    }
});

app.post('/api/admin/push-event', authMiddleware, async (req, res) => {
    try {
        const eventType = String(req.body?.eventType || '').trim().toLowerCase();
        const title = String(req.body?.title || '').trim();
        const body = String(req.body?.body || '').trim();
        const clickUrl = String(req.body?.clickUrl || '/admin-dashboard.html').trim() || '/admin-dashboard.html';
        const meta = req.body?.meta || {};

        if (!eventType || !title || !body) {
            return res.status(400).json({ ok: false, error: 'eventType, title, and body are required' });
        }

        const allowedEventTypes = new Set([
            'new_user_registration',
            'new_agent_registration'
        ]);
        if (!allowedEventTypes.has(eventType)) {
            return res.status(400).json({ ok: false, error: 'Unsupported admin push event type' });
        }

        const allUsersSnap = await adminDb.collection('users').get();
        const adminRecipients = [];
        const tokenOwners = [];

        allUsersSnap.docs.forEach((docSnap) => {
            const data = docSnap.data() || {};
            const role = normalizeRole(data.role);
            const status = String(data.status || '').toLowerCase();
            if (!['admin', 'super_admin'].includes(role)) return;
            if (status && status !== 'active') return;
            const email = safeLowerEmail(data.email);
            adminRecipients.push(email || docSnap.id);
            const tokens = Array.isArray(data.fcmTokens) ? data.fcmTokens : [];
            tokens.forEach((token) => {
                const cleaned = String(token || '').trim();
                if (cleaned) {
                    tokenOwners.push({ email, userId: docSnap.id, token: cleaned });
                }
            });
        });

        const uniqueMap = new Map();
        tokenOwners.forEach((entry) => {
            if (!uniqueMap.has(entry.token)) uniqueMap.set(entry.token, entry);
        });
        const uniqueOwners = Array.from(uniqueMap.values());
        const tokens = uniqueOwners.map((entry) => entry.token);

        if (tokens.length === 0) {
            return res.json({
                ok: true,
                sent: 0,
                reason: 'no-device-tokens',
                debug: { adminRecipients }
            });
        }

        const response = await admin.messaging().sendEachForMulticast({
            tokens,
            data: {
                eventType,
                clickUrl,
                title,
                body,
                meta: JSON.stringify(meta || {})
            },
            webpush: {
                fcmOptions: { link: clickUrl }
            }
        });

        const badTokens = [];
        response.responses.forEach((r, idx) => {
            if (r.success) return;
            const code = String(r.error?.code || '');
            if (
                code.includes('registration-token-not-registered') ||
                code.includes('invalid-argument')
            ) {
                badTokens.push(uniqueOwners[idx]);
            }
        });

        for (const bad of badTokens) {
            if (!bad?.userId) continue;
            try {
                const ref = adminDb.collection('users').doc(bad.userId);
                const snap = await ref.get();
                if (!snap.exists) continue;
                const arr = Array.isArray(snap.data()?.fcmTokens) ? snap.data().fcmTokens : [];
                const next = arr.filter((t) => String(t || '').trim() !== bad.token);
                await ref.set({ fcmTokens: next, fcmTokenPrunedAt: admin.firestore.FieldValue.serverTimestamp() }, { merge: true });
            } catch (_) {}
        }

        return res.json({
            ok: true,
            sent: response.successCount,
            failed: response.failureCount,
            debug: {
                adminRecipients,
                tokenOwners: uniqueOwners.map((entry) => entry.email || entry.userId)
            }
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send admin push event')
        });
    }
});

app.post('/api/user/push-event', authMiddleware, async (req, res) => {
    try {
        const recipientUserId = String(req.body?.recipientUserId || '').trim();
        const recipientEmail = safeLowerEmail(req.body?.recipientEmail);
        const eventType = String(req.body?.eventType || '').trim().toLowerCase();
        const title = String(req.body?.title || '').trim();
        const body = String(req.body?.body || '').trim();
        const clickUrl = String(req.body?.clickUrl || '/dashboard.html').trim() || '/dashboard.html';
        const meta = req.body?.meta || {};

        if ((!recipientUserId && !recipientEmail) || !eventType || !title || !body) {
            return res.status(400).json({ ok: false, error: 'recipient, eventType, title, and body are required' });
        }

        const allowedEventTypes = new Set([
            'agent_registration_approved'
        ]);
        if (!allowedEventTypes.has(eventType)) {
            return res.status(400).json({ ok: false, error: 'Unsupported user push event type' });
        }

        let targetDoc = null;
        if (recipientUserId) {
            const byIdSnap = await adminDb.collection('users').doc(recipientUserId).get();
            if (byIdSnap.exists) {
                targetDoc = { id: byIdSnap.id, ...(byIdSnap.data() || {}) };
            }
        }
        if (!targetDoc && recipientEmail) {
            const byEmailMap = await resolveUserMapByEmail([recipientEmail]);
            targetDoc = byEmailMap.get(recipientEmail) || null;
        }

        if (!targetDoc) {
            return res.status(404).json({ ok: false, error: 'Recipient user not found' });
        }

        const role = normalizeRole(targetDoc.role);
        const status = String(targetDoc.status || '').toLowerCase();
        if (role === 'super_admin' || status === 'pending' || status === 'deactivated') {
            return res.json({ ok: true, sent: 0, reason: 'recipient-not-eligible' });
        }

        const tokenOwners = [];
        const tokens = Array.isArray(targetDoc.fcmTokens) ? targetDoc.fcmTokens : [];
        tokens.forEach((token) => {
            const cleaned = String(token || '').trim();
            if (cleaned) {
                tokenOwners.push({
                    email: safeLowerEmail(targetDoc.email),
                    userId: targetDoc.id || recipientUserId,
                    token: cleaned
                });
            }
        });

        const uniqueMap = new Map();
        tokenOwners.forEach((entry) => {
            if (!uniqueMap.has(entry.token)) uniqueMap.set(entry.token, entry);
        });
        const uniqueOwners = Array.from(uniqueMap.values());
        const pushTokens = uniqueOwners.map((entry) => entry.token);

        if (pushTokens.length === 0) {
            return res.json({
                ok: true,
                sent: 0,
                reason: 'no-device-tokens',
                debug: {
                    recipientUserId: targetDoc.id || recipientUserId,
                    recipientEmail: safeLowerEmail(targetDoc.email) || recipientEmail
                }
            });
        }

        const response = await admin.messaging().sendEachForMulticast({
            tokens: pushTokens,
            data: {
                eventType,
                clickUrl,
                title,
                body,
                meta: JSON.stringify(meta || {})
            },
            webpush: {
                fcmOptions: { link: clickUrl }
            }
        });

        const badTokens = [];
        response.responses.forEach((r, idx) => {
            if (r.success) return;
            const code = String(r.error?.code || '');
            if (
                code.includes('registration-token-not-registered') ||
                code.includes('invalid-argument')
            ) {
                badTokens.push(uniqueOwners[idx]);
            }
        });

        for (const bad of badTokens) {
            if (!bad?.userId) continue;
            try {
                const ref = adminDb.collection('users').doc(bad.userId);
                const snap = await ref.get();
                if (!snap.exists) continue;
                const arr = Array.isArray(snap.data()?.fcmTokens) ? snap.data().fcmTokens : [];
                const next = arr.filter((t) => String(t || '').trim() !== bad.token);
                await ref.set({ fcmTokens: next, fcmTokenPrunedAt: admin.firestore.FieldValue.serverTimestamp() }, { merge: true });
            } catch (_) {}
        }

        return res.json({
            ok: true,
            sent: response.successCount,
            failed: response.failureCount,
            debug: {
                recipientUserId: targetDoc.id || recipientUserId,
                recipientEmail: safeLowerEmail(targetDoc.email) || recipientEmail
            }
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send user push event')
        });
    }
});

let lastScheduledDueKey = '';

async function tickScheduledReportSender() {
    if (!ENABLE_SCHEDULED_REPORT_SENDER) return;
    const settingsSnap = await adminDb.collection('settings').doc('system').get();
    const settings = settingsSnap.exists ? (settingsSnap.data() || {}) : {};
    const scheduled = settings.scheduledReportEmail || {};
    if (scheduled.enabled !== true) return;

    const sendTime = String(scheduled.sendTime || '08:00').trim() || '08:00';
    const now = new Date();
    const lagosDateKey = getLagosDateKey(now);
    const dueKey = `${lagosDateKey}:${sendTime}`;
    if (!isScheduledTimeDue(now, sendTime) || dueKey === lastScheduledDueKey) return;

    lastScheduledDueKey = dueKey;
    const autoScope = resolveAutoScheduledScope(now);
    if (autoScope.mode === 'skip') return;
    try {
        if (autoScope.mode === 'date_range') {
            await sendManualReport({
                rangeStartDateKey: autoScope.rangeStartDateKey,
                rangeEndDateKey: autoScope.rangeEndDateKey,
                trigger: 'auto_weekend_rollup'
            });
            return;
        }

        await sendScheduledReportForDate({
            reportDateKey: autoScope.reportDateKey,
            trigger: 'auto'
        });
    } catch (err) {
        console.error('[scheduled-report] auto send failed:', err?.message || err);
    }
}

setInterval(() => {
    tickScheduledReportSender().catch((err) => {
        console.error('[scheduled-report] tick failed:', err?.message || err);
    });
}, 30000);

tickScheduledReportSender().catch((err) => {
    console.error('[scheduled-report] initial tick failed:', err?.message || err);
});

app.listen(port);
