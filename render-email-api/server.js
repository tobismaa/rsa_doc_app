require('dotenv').config();

const express = require('express');
const cors = require('cors');
const admin = require('firebase-admin');
const sgMail = require('@sendgrid/mail');

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
    const v = String(value || '').trim().toLowerCase();
    if (v === 'viewer') return 'reviewer';
    return v;
}

function safeLowerEmail(value) {
    return normalizeEmail(value);
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
        const chatRef = adminDb.collection('applicationChats').doc(submissionId);
        const [subSnap, chatSnap] = await Promise.all([subRef.get(), chatRef.get()]);
        if (!subSnap.exists) {
            return res.status(404).json({ ok: false, error: 'Submission not found' });
        }
        const sub = subSnap.data() || {};
        const chat = chatSnap.exists ? (chatSnap.data() || {}) : {};

        const participantEmails = Array.from(new Set([
            safeLowerEmail(sub.uploadedBy),
            safeLowerEmail(sub.assignedTo),
            safeLowerEmail(sub.reviewedBy),
            safeLowerEmail(sub.assignedToRSA),
            safeLowerEmail(sub.assignedToPayment),
            ...(Array.isArray(chat.participants) ? chat.participants.map(safeLowerEmail) : []),
            ...(Array.isArray(chat.adminParticipants) ? chat.adminParticipants.map(safeLowerEmail) : [])
        ].filter(Boolean)));

        const senderCanAccess = participantEmails.includes(senderEmail);
        if (!senderCanAccess) {
            const senderUserSnap = await adminDb.collection('users').where('email', '==', senderEmail).limit(1).get();
            const senderRole = !senderUserSnap.empty ? normalizeRole(senderUserSnap.docs[0].data()?.role) : '';
            if (!['admin', 'super_admin'].includes(senderRole)) {
                return res.status(403).json({ ok: false, error: 'Not permitted for this submission' });
            }
        }

        const recipientEmailsRaw = participantEmails.filter((e) => e !== senderEmail);
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

        // Exclude super admin from chat notifications/recipients.
        recipientEmails = recipientEmails.filter((email) => {
            const user = usersByEmail.get(email);
            const role = normalizeRole(user?.role);
            return role !== 'super_admin';
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
            notification: { title, body },
            data: {
                submissionId,
                clickUrl
            },
            webpush: {
                fcmOptions: { link: clickUrl },
                notification: {
                    title,
                    body,
                    icon: '/favicon.svg',
                    badge: '/favicon.svg',
                    data: { url: clickUrl }
                }
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

app.listen(port, () => {
    console.log(`RSA email API running on port ${port}`);
});
