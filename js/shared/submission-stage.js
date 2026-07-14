function getTimestampMillis(value) {
    if (!value) return 0;
    try {
        if (typeof value.toMillis === 'function') return value.toMillis();
        if (typeof value.toDate === 'function') return value.toDate().getTime();
        if (typeof value.seconds === 'number') return value.seconds * 1000;
    } catch (_) {}
    const parsed = new Date(value).getTime();
    return Number.isFinite(parsed) ? parsed : 0;
}

function pickTimestamp(...values) {
    for (const value of values) {
        if (getTimestampMillis(value) > 0) return value;
    }
    return null;
}

function getSubmissionOriginalUploadAt(submission = {}) {
    return pickTimestamp(
        submission?.originalUploadedAt,
        submission?.firstUploadedAt,
        submission?.initialUploadedAt,
        submission?.submittedAt,
        submission?.createdAt,
        submission?.uploadedAt
    );
}

function getSubmissionDraftEntryAt(submission = {}) {
    return pickTimestamp(submission?.draftSavedAt, submission?.updatedAt, submission?.uploadedAt, submission?.createdAt);
}

function getSubmissionReviewEntryAt(submission = {}) {
    return pickTimestamp(submission?.reuploadedAt, submission?.uploadedAt, submission?.submittedAt, submission?.createdAt, submission?.updatedAt);
}

function getSubmissionApprovalEntryAt(submission = {}) {
    return pickTimestamp(submission?.reviewedAt, submission?.approvedAt, submission?.statusUpdatedAt, submission?.updatedAt);
}

function getSubmissionRejectionEntryAt(submission = {}) {
    return pickTimestamp(
        submission?.latestRejectedAt,
        submission?.rejectedAt,
        submission?.previousRejectedAt,
        submission?.reviewedAt,
        submission?.statusUpdatedAt,
        submission?.updatedAt
    );
}

function getSubmissionRsaEntryAt(submission = {}) {
    return pickTimestamp(
        submission?.rsaAssignedAt,
        submission?.reviewedAt,
        submission?.approvedAt,
        submission?.reuploadedAt,
        submission?.uploadedAt,
        submission?.submittedAt,
        submission?.createdAt,
        submission?.statusUpdatedAt,
        submission?.updatedAt
    );
}

function getSubmissionFinalSubmissionEntryAt(submission = {}) {
    return pickTimestamp(submission?.finalSubmittedAt, submission?.rsaSubmittedAt, submission?.statusUpdatedAt, submission?.updatedAt);
}

function getSubmissionPaymentEntryAt(submission = {}) {
    return pickTimestamp(submission?.paymentAssignedAt, submission?.finalSubmittedAt, submission?.rsaSubmittedAt, submission?.statusUpdatedAt, submission?.updatedAt);
}

function getSubmissionPaidEntryAt(submission = {}) {
    return pickTimestamp(submission?.paidAt, submission?.statusUpdatedAt, submission?.updatedAt);
}

function getSubmissionClearedEntryAt(submission = {}) {
    return pickTimestamp(submission?.clearedAt, submission?.statusUpdatedAt, submission?.updatedAt);
}

function getSubmissionCurrentStageEntryAt(submission = {}) {
    const status = String(submission?.status || '').trim().toLowerCase();
    if (status === 'draft') return getSubmissionDraftEntryAt(submission);
    if (['pending', 'submitted', 'resubmitted'].includes(status)) return getSubmissionReviewEntryAt(submission);
    if (['rejected', 'rejected_by_reviewer', 'rejected_by_rsa'].includes(status)) return getSubmissionRejectionEntryAt(submission);
    if (['approved', 'processing_to_pfa'].includes(status)) return getSubmissionRsaEntryAt(submission);
    if (['sent_to_pfa', 'rsa_submitted'].includes(status) || submission?.finalSubmitted === true || submission?.rsaSubmitted === true) {
        return getSubmissionPaymentEntryAt(submission);
    }
    if (status === 'paid') return getSubmissionPaidEntryAt(submission);
    if (status === 'cleared') return getSubmissionClearedEntryAt(submission);
    return pickTimestamp(submission?.statusUpdatedAt, submission?.updatedAt, submission?.uploadedAt, submission?.createdAt);
}

export {
    getTimestampMillis,
    getSubmissionDraftEntryAt,
    getSubmissionReviewEntryAt,
    getSubmissionApprovalEntryAt,
    getSubmissionRejectionEntryAt,
    getSubmissionRsaEntryAt,
    getSubmissionFinalSubmissionEntryAt,
    getSubmissionPaymentEntryAt,
    getSubmissionPaidEntryAt,
    getSubmissionClearedEntryAt,
    getSubmissionOriginalUploadAt,
    getSubmissionCurrentStageEntryAt
};
