import { auth, db } from './firebase-config.js';
import { EMAIL_API_BASE_URL } from './email-api-config.js';
import {
    doc,
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

async function queueEmailNotification({ eventKey, to, subject, textLines, htmlLines, meta = {} }) {
    const recipient = normalizeEmail(to);
    const dedupeKey = String(eventKey || '').trim();
    if (!recipient || !dedupeKey) return { queued: false, reason: 'missing-recipient-or-key' };

    const safeSubject = String(subject || 'RSA Portal Notification').trim() || 'RSA Portal Notification';
    const text = (Array.isArray(textLines) ? textLines : [])
        .map(line => String(line || '').trim())
        .filter(Boolean)
        .join('\n');
    const htmlBody = (Array.isArray(htmlLines) ? htmlLines : [])
        .map(line => `<p style="margin:0 0 10px;">${escapeHtml(line)}</p>`)
        .join('');
    const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.5;color:#0f172a;">${htmlBody}</div>`;

    let queued = false;
    const guardRef = doc(db, 'emailNotificationEvents', dedupeKey);

    await runTransaction(db, async (tx) => {
        const guardSnap = await tx.get(guardRef);
        if (guardSnap.exists()) return;

        tx.set(guardRef, {
            eventKey: dedupeKey,
            recipient,
            subject: safeSubject,
            status: 'queued',
            createdAt: serverTimestamp(),
            meta
        });

        queued = true;
    });

    if (!queued) return { queued: false, reason: 'duplicate' };

    const apiBaseUrl = String(EMAIL_API_BASE_URL || '').trim().replace(/\/+$/, '');
    if (!apiBaseUrl || apiBaseUrl.includes('YOUR-RENDER-URL')) {
        await updateDoc(guardRef, {
            status: 'failed',
            lastError: 'Email API URL is not configured',
            failedAt: serverTimestamp()
        });
        return { queued: true, sent: false, reason: 'email-api-not-configured' };
    }

    try {
        const idToken = auth.currentUser ? await auth.currentUser.getIdToken() : '';
        const response = await fetch(`${apiBaseUrl}/api/send-alert`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                ...(idToken ? { 'Authorization': `Bearer ${idToken}` } : {})
            },
            body: JSON.stringify({
                eventKey: dedupeKey,
                to: recipient,
                subject: safeSubject,
                text,
                html,
                meta
            })
        });

        const result = await response.json().catch(() => ({}));
        if (!response.ok || !result.ok) {
            throw new Error(result.error || `Email API returned ${response.status}`);
        }

        try {
            await updateDoc(guardRef, {
                status: 'sent',
                sentAt: serverTimestamp(),
                provider: 'sendgrid',
                providerMessageId: String(result.messageId || '')
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
        meta: { type: 'submission_assigned_viewer', submissionId }
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
        meta: { type: 'submission_rejected_uploader', submissionId }
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
        meta: { type: 'submission_approved_rsa', submissionId }
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
        meta: { type: 'submission_approved_uploader', submissionId }
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
        meta: { type: 'submission_final_uploader', submissionId }
    });
}
