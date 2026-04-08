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

const sendgridApiKey = String(process.env.SENDGRID_API_KEY || '').trim();
if (!sendgridApiKey) {
    throw new Error('SENDGRID_API_KEY is required');
}
sgMail.setApiKey(sendgridApiKey);

const defaultFromEmail = normalizeEmail(process.env.SENDGRID_FROM_EMAIL);
if (!defaultFromEmail) {
    throw new Error('SENDGRID_FROM_EMAIL is required and must be a valid email');
}

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

app.post('/api/send-alert', authMiddleware, async (req, res) => {
    try {
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

        const headerMessageId = sendResult?.headers?.['x-message-id'] || '';
        return res.json({
            ok: true,
            messageId: String(headerMessageId || '')
        });
    } catch (err) {
        return res.status(500).json({
            ok: false,
            error: String(err?.message || 'Failed to send alert email')
        });
    }
});

app.listen(port, () => {
    console.log(`RSA email API running on port ${port}`);
});
