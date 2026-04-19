import { db } from './firebase-config.js';
import {
    EMAILJS_PUBLIC_KEY,
    EMAILJS_SERVICE_ID,
    EMAILJS_TEMPLATE_ID
} from './emailjs-config.js';
import {
    doc,
    increment,
    runTransaction,
    serverTimestamp,
    updateDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

function normalizeEmail(email) {
    const trimmed = String(email || '').trim().toLowerCase();
    return trimmed.includes('@') ? trimmed : '';
}

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function titleCaseWord(word) {
    const w = String(word || '').trim();
    if (!w) return '';
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
}

function getFirstNameFromEmail(email) {
    const local = String(email || '').split('@')[0] || '';
    const firstChunk = local.split(/[._\-\s]+/)[0] || '';
    return titleCaseWord(firstChunk);
}

function getRecipientFirstName({ recipientEmail, preferredName }) {
    const direct = titleCaseWord(preferredName);
    if (direct) return direct;
    const fromEmail = getFirstNameFromEmail(recipientEmail);
    return fromEmail || 'User';
}

const PORTAL_URL = window.location.origin;
const ALERT_REPLY_TO = 'alert@cmbankrsa.com';
// Optional: set to a hosted GIF URL if you want anime-style motion in emails.
// Keep empty to disable.
const ANIME_GIF_URL = '';

function isEmailJsConfigured() {
    const parts = [EMAILJS_PUBLIC_KEY, EMAILJS_SERVICE_ID, EMAILJS_TEMPLATE_ID]
        .map(v => String(v || '').trim());
    return parts.every(v => v && !v.includes('REPLACE_WITH_'));
}

async function sendEmailViaEmailJs({ eventKey, to, subject, text, html, meta = {} }) {
    const templateParams = {
        to_email: to,
        reply_to: ALERT_REPLY_TO,
        from_email: ALERT_REPLY_TO,
        from_name: 'CMBank RSA Alerts',
        email_subject: subject,
        email_message: text,
        email_html: html,
        customer_name: meta.customerName || '',
        submission_id: meta.submissionId || '',
        notification_type: meta.type || '',
        event_key: eventKey || '',
        subject,
        message: text,
        html_content: html
    };

    const payload = {
        service_id: EMAILJS_SERVICE_ID,
        template_id: EMAILJS_TEMPLATE_ID,
        user_id: EMAILJS_PUBLIC_KEY,
        template_params: templateParams
    };

    for (let attempt = 1; attempt <= 2; attempt += 1) {
        try {
            const response = await fetch('https://api.emailjs.com/api/v1.0/email/send', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const body = await response.text().catch(() => '');
                throw new Error(body || `EmailJS returned ${response.status}`);
            }

            return;
        } catch (err) {
            const msg = String(err?.message || err || '');
            const networkError = msg.toLowerCase().includes('failed to fetch');
            if (attempt < 2 && networkError) {
                await new Promise((resolve) => setTimeout(resolve, 1200));
                continue;
            }
            throw err;
        }
    }
}

async function queueEmailNotification({ eventKey, to, subject, textLines, htmlLines, meta = {} }) {
    const recipient = normalizeEmail(to);
    const dedupeKey = String(eventKey || '').trim();
    if (!recipient || !dedupeKey) return { queued: false, reason: 'missing-recipient-or-key' };

    const safeSubject = String(subject || 'RSA Portal Notification').trim() || 'RSA Portal Notification';
    const firstName = getRecipientFirstName({
        recipientEmail: recipient,
        preferredName: meta?.recipientFirstName || ''
    });
    const coreTextLines = (Array.isArray(textLines) ? textLines : [])
        .map(line => String(line || '').trim())
        .filter(Boolean);
    const text = [`Dear ${firstName},`, '', ...coreTextLines, '', `Portal: ${PORTAL_URL}`, '', 'Best regards,', 'Admin']
        .join('\n');
    const coreHtmlLines = (Array.isArray(htmlLines) ? htmlLines : [])
        .map(line => `<p style="margin:0 0 12px;color:#1f2937;font-size:15px;line-height:1.6;">${escapeHtml(line)}</p>`)
        .join('');
    const animeBlock = ANIME_GIF_URL
        ? `<div style="margin:0 0 16px;text-align:center;">
             <img src="${escapeHtml(ANIME_GIF_URL)}" alt="Notification Visual" style="max-width:100%;width:100%;border-radius:10px;border:1px solid #e5e7eb;">
           </div>`
        : '';
    const html = `
<div style="margin:0;padding:20px;background:#f4f7fb;font-family:Segoe UI,Arial,sans-serif;">
  <div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #dbe4f0;border-radius:14px;overflow:hidden;">
    <div style="background:linear-gradient(135deg,#003366 0%,#0b5cab 100%);padding:18px 22px;">
      <div style="color:#e2ecff;font-size:12px;letter-spacing:0.08em;text-transform:uppercase;">CMBank RSA</div>
      <h2 style="margin:6px 0 0;color:#ffffff;font-size:21px;line-height:1.3;">${escapeHtml(safeSubject)}</h2>
    </div>
    <div style="padding:22px;">
      ${animeBlock}
      <p style="margin:0 0 12px;color:#1f2937;font-size:15px;line-height:1.6;">Dear ${escapeHtml(firstName)},</p>
      ${coreHtmlLines}
      <div style="margin-top:16px;">
        <a href="${PORTAL_URL}" style="display:inline-block;background:#003366;color:#ffffff;text-decoration:none;padding:11px 16px;border-radius:8px;font-weight:600;">
          Open CMBank RSA Portal
        </a>
      </div>
      <p style="margin:14px 0 0;color:#1f2937;font-size:15px;line-height:1.6;">Best regards,<br>Admin</p>
      <p style="margin:14px 0 0;color:#475569;font-size:12px;">
        Direct link: <a href="${PORTAL_URL}" style="color:#0b5cab;text-decoration:underline;">${PORTAL_URL}</a>
      </p>
    </div>
    <div style="padding:12px 22px;background:#f8fafc;border-top:1px solid #e2e8f0;color:#64748b;font-size:12px;">
      Automated notification from CMBank RSA Document Portal.
    </div>
  </div>
</div>`;

    let queued = false;
    const guardRef = doc(db, 'emailNotificationEvents', dedupeKey);

    await runTransaction(db, async (tx) => {
        const guardSnap = await tx.get(guardRef);
        if (guardSnap.exists()) {
            const prev = guardSnap.data() || {};
            const prevStatus = String(prev.status || '').toLowerCase();
            if (prevStatus === 'sent') return;

            tx.set(guardRef, {
                status: 'queued',
                retryCount: increment(1),
                retriedAt: serverTimestamp(),
                lastError: ''
            }, { merge: true });

            queued = true;
            return;
        }

        tx.set(guardRef, {
            eventKey: dedupeKey,
            recipient,
            subject: safeSubject,
            status: 'queued',
            retryCount: 0,
            createdAt: serverTimestamp(),
            meta
        });

        queued = true;
    });

    if (!queued) return { queued: false, reason: 'duplicate' };

    if (!isEmailJsConfigured()) {
        await updateDoc(guardRef, {
            status: 'failed',
            lastError: 'EmailJS config is not set',
            failedAt: serverTimestamp()
        });
        return { queued: true, sent: false, reason: 'emailjs-not-configured' };
    }

    try {
        await sendEmailViaEmailJs({
            eventKey: dedupeKey,
            to: recipient,
            subject: safeSubject,
            text,
            html,
            meta
        });

        try {
            await updateDoc(guardRef, {
                status: 'sent',
                sentAt: serverTimestamp(),
                provider: 'emailjs'
            });
        } catch (_) {}

        return { queued: true, sent: true };
    } catch (err) {
        await updateDoc(guardRef, {
            status: 'failed',
            lastError: String(err?.message || err || 'unknown-error'),
            failedAt: serverTimestamp()
        });
        return { queued: true, sent: false, reason: 'send-failed' };
    }
}

export async function queueViewerAssignmentEmail({ submissionId, viewerEmail, customerName, uploaderEmail }) {
    return queueEmailNotification({
        eventKey: `submission-assigned-viewer-${submissionId}`,
        to: viewerEmail,
        subject: `New Submission Assigned: ${customerName || 'Customer'}`,
        textLines: [
            'A new submission has been assigned to you for review.',
            `Customer: ${customerName || 'N/A'}`,
            `Uploaded by: ${uploaderEmail || 'N/A'}`,
      'Please login to the Reviewer Dashboard to process it.'
        ],
        htmlLines: [
            'A new submission has been assigned to you for review.',
            `Customer: ${customerName || 'N/A'}`,
            `Uploaded by: ${uploaderEmail || 'N/A'}`,
      'Please login to the Reviewer Dashboard to process it.'
        ],
        meta: { type: 'submission_assigned_viewer', submissionId, customerName }
    });
}

export async function queueUploaderRejectedEmail({ submissionId, uploaderEmail, customerName, reviewerEmail, rejectionReason }) {
    return queueEmailNotification({
        eventKey: `submission-rejected-uploader-${submissionId}`,
        to: uploaderEmail,
        subject: `Submission Rejected: ${customerName || 'Customer'}`,
        textLines: [
            'Your submission has been rejected.',
            `Customer: ${customerName || 'N/A'}`,
            `Rejected by: ${reviewerEmail || 'N/A'}`,
            `Reason: ${rejectionReason || 'No reason provided'}`,
            'Please login to the Uploader Dashboard to correct and re-submit.'
        ],
        htmlLines: [
            'Your submission has been rejected.',
            `Customer: ${customerName || 'N/A'}`,
            `Rejected by: ${reviewerEmail || 'N/A'}`,
            `Reason: ${rejectionReason || 'No reason provided'}`,
            'Please login to the Uploader Dashboard to correct and re-submit.'
        ],
        meta: { type: 'submission_rejected_uploader', submissionId, customerName }
    });
}

export async function queueRsaApprovalEmail({ submissionId, rsaEmail, customerName, reviewerEmail, uploaderEmail }) {
    return queueEmailNotification({
        eventKey: `submission-approved-rsa-${submissionId}`,
        to: rsaEmail,
        subject: `Approved Submission Assigned: ${customerName || 'Customer'}`,
        textLines: [
            'A submission has been approved and assigned to you for RSA processing.',
            `Customer: ${customerName || 'N/A'}`,
            `Reviewed by: ${reviewerEmail || 'N/A'}`,
            `Uploaded by: ${uploaderEmail || 'N/A'}`,
            'Please login to the RSA Dashboard to continue.'
        ],
        htmlLines: [
            'A submission has been approved and assigned to you for RSA processing.',
            `Customer: ${customerName || 'N/A'}`,
            `Reviewed by: ${reviewerEmail || 'N/A'}`,
            `Uploaded by: ${uploaderEmail || 'N/A'}`,
            'Please login to the RSA Dashboard to continue.'
        ],
        meta: { type: 'submission_approved_rsa', submissionId, customerName }
    });
}

export async function queueUploaderApprovedEmail({ submissionId, uploaderEmail, customerName, reviewerEmail, rsaEmail }) {
    return queueEmailNotification({
        eventKey: `submission-approved-uploader-${submissionId}`,
        to: uploaderEmail,
        subject: `Submission Approved: ${customerName || 'Customer'}`,
        textLines: [
            'Your submission has been approved.',
            `Customer: ${customerName || 'N/A'}`,
            `Approved by: ${reviewerEmail || 'N/A'}`,
            rsaEmail ? `Assigned to RSA: ${rsaEmail}` : 'Assigned to RSA: Pending',
            'You can track progress in the Uploader Dashboard.'
        ],
        htmlLines: [
            'Your submission has been approved.',
            `Customer: ${customerName || 'N/A'}`,
            `Approved by: ${reviewerEmail || 'N/A'}`,
            rsaEmail ? `Assigned to RSA: ${rsaEmail}` : 'Assigned to RSA: Pending',
            'You can track progress in the Uploader Dashboard.'
        ],
        meta: { type: 'submission_approved_uploader', submissionId, customerName }
    });
}

export async function queueUploaderFinalSubmissionEmail({ submissionId, uploaderEmail, customerName, rsaEmail }) {
    return queueEmailNotification({
        eventKey: `submission-final-uploader-${submissionId}`,
        to: uploaderEmail,
        subject: `Final Approval Completed: ${customerName || 'Customer'}`,
        textLines: [
            'RSA processing has been completed for your submission.',
            `Customer: ${customerName || 'N/A'}`,
            `Finalized by RSA: ${rsaEmail || 'N/A'}`,
            'Please login to the Uploader Dashboard to view the final status.'
        ],
        htmlLines: [
            'RSA processing has been completed for your submission.',
            `Customer: ${customerName || 'N/A'}`,
            `Finalized by RSA: ${rsaEmail || 'N/A'}`,
            'Please login to the Uploader Dashboard to view the final status.'
        ],
        meta: { type: 'submission_final_uploader', submissionId, customerName }
    });
}
