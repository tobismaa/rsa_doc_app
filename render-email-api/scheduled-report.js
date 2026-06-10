const ExcelJS = require('exceljs');
const REPORT_INCEPTION_START_DATE = '1900-01-01';
const REPORT_INCEPTION_LABEL = 'From Inception';

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

function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
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

function getLagosDateKey(date = new Date()) {
    return lagosDateKeyFormatter.format(date);
}

function getPreviousDateKeyInLagos(baseDate = new Date()) {
    const prior = new Date(baseDate);
    prior.setDate(prior.getDate() - 1);
    return getLagosDateKey(prior);
}

function pickTimestamp(...values) {
    for (const value of values) {
        if (tsToMillis(value) > 0) return value;
    }
    return null;
}

function getSubmissionReviewEntryAt(sub = {}) {
    return pickTimestamp(sub.reuploadedAt, sub.uploadedAt, sub.submittedAt, sub.createdAt, sub.updatedAt);
}

function getSubmissionApprovalEntryAt(sub = {}) {
    return pickTimestamp(sub.reviewedAt, sub.approvedAt, sub.statusUpdatedAt, sub.updatedAt);
}

function getSubmissionRsaEntryAt(sub = {}) {
    return pickTimestamp(
        sub.rsaAssignedAt,
        sub.reviewedAt,
        sub.approvedAt,
        sub.reuploadedAt,
        sub.uploadedAt,
        sub.submittedAt,
        sub.createdAt,
        sub.statusUpdatedAt,
        sub.updatedAt
    );
}

function getSubmissionFinalSubmissionEntryAt(sub = {}) {
    return pickTimestamp(sub.finalSubmittedAt, sub.rsaSubmittedAt, sub.statusUpdatedAt, sub.updatedAt);
}

function getSubmissionPaymentEntryAt(sub = {}) {
    return pickTimestamp(sub.paymentAssignedAt, sub.finalSubmittedAt, sub.rsaSubmittedAt, sub.statusUpdatedAt, sub.updatedAt);
}

function getSubmissionPaidEntryAt(sub = {}) {
    return pickTimestamp(sub.paidAt, sub.statusUpdatedAt, sub.updatedAt);
}

function getSubmissionClearedEntryAt(sub = {}) {
    return pickTimestamp(sub.clearedAt, sub.statusUpdatedAt, sub.updatedAt);
}

function formatDate(value) {
    const ms = tsToMillis(value);
    if (!ms) return '-';
    return lagosDateTimeFormatter.format(new Date(ms));
}

function getDateKey(value) {
    const ms = tsToMillis(value);
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
    const num = Number(value || 0);
    return Math.max(0, Math.floor(num / 1000) * 1000);
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

function getRejectionReason(sub = {}) {
    return String(sub?.latestRejectionReason || sub?.previousRejectionReason || sub?.comment || '').trim();
}

function getRejectionOfficerName(sub = {}, usersByEmail) {
    if (getRejectionCount(sub) <= 0 && !getRejectionReason(sub)) return '';
    const rejectedBy = String(sub?.latestRejectedBy || sub?.previousRejectedBy || '').trim();
    if (rejectedBy) return getUserDisplayNameByEmail(usersByEmail, rejectedBy);
    const stage = String(sub?.latestRejectedStage || '').trim().toLowerCase();
    if (stage === 'rsa') return getUserDisplayNameByEmail(usersByEmail, sub?.assignedToRSA || '');
    if (stage === 'payment') return getUserDisplayNameByEmail(usersByEmail, sub?.assignedToPayment || '');
    return getUserDisplayNameByEmail(usersByEmail, sub?.reviewedBy || sub?.assignedTo || '');
}

function getRejectionCount(sub = {}) {
    const history = Array.isArray(sub?.rejectionHistory)
        ? sub.rejectionHistory.filter((entry) => {
            if (typeof entry === 'string') return String(entry).trim();
            return String(entry?.reason || '').trim();
        })
        : [];
    if (history.length) return history.length;
    return getRejectionReason(sub) ? 1 : 0;
}

function formatMoneyForSheet(value) {
    const num = Number(value || 0);
    return num ? num : '';
}

function getUserDisplayNameByEmail(usersByEmail, email = '') {
    const normalized = normalizeEmail(email);
    if (!normalized) return 'Unassigned';
    return String(usersByEmail.get(normalized)?.fullName || normalized).trim();
}

function compareGroupedRows(a = {}, b = {}) {
    const ownerA = String(a.owner || '').toLowerCase();
    const ownerB = String(b.owner || '').toLowerCase();
    if (ownerA !== ownerB) return ownerA.localeCompare(ownerB);

    const dateA = String(a.reportDate || '').toLowerCase();
    const dateB = String(b.reportDate || '').toLowerCase();
    if (dateA !== dateB) return dateA.localeCompare(dateB);

    const assignedA = String(a.assignedAt || a.uploadedAt || '').toLowerCase();
    const assignedB = String(b.assignedAt || b.uploadedAt || '').toLowerCase();
    if (assignedA !== assignedB) return assignedA.localeCompare(assignedB);

    const stageA = String(a.stageTime || a.paidAt || a.clearedAt || '').toLowerCase();
    const stageB = String(b.stageTime || b.paidAt || b.clearedAt || '').toLowerCase();
    if (stageA !== stageB) return stageA.localeCompare(stageB);

    return String(a.customerName || '').toLowerCase().localeCompare(String(b.customerName || '').toLowerCase());
}

function buildUploaderSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records.map((sub) => ({
        owner: getUserDisplayNameByEmail(usersByEmail, sub.uploadedBy),
        ...(includeDateColumn ? { reportDate: getDateKey(getSubmissionReviewEntryAt(sub)) || '-' } : {}),
        customerName: sub.customerName || '',
        rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
        rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
        commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
        status: String(sub.status || '').replace(/_/g, ' '),
        uploadedAt: formatDate(getSubmissionReviewEntryAt(sub)),
        rejectionReason: getRejectionReason(sub),
        rejectionOfficer: getRejectionOfficerName(sub, usersByEmail),
        rejectionCount: getRejectionCount(sub)
    }));
}

function buildReviewerSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedTo))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(usersByEmail, sub.assignedTo),
            ...(includeDateColumn ? { reportDate: getDateKey(getSubmissionReviewEntryAt(sub)) || '-' } : {}),
            customerName: sub.customerName || '',
            uploaderName: getUserDisplayNameByEmail(usersByEmail, sub.uploadedBy),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionReviewEntryAt(sub)),
            rejectionReason: getRejectionReason(sub),
            rejectionCount: getRejectionCount(sub)
        }));
}

function buildRsaSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedToRSA))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(usersByEmail, sub.assignedToRSA),
            ...(includeDateColumn ? { reportDate: getDateKey(getSubmissionRsaEntryAt(sub)) || '-' } : {}),
            customerName: sub.customerName || '',
            uploaderName: getUserDisplayNameByEmail(usersByEmail, sub.uploadedBy),
            reviewerName: getUserDisplayNameByEmail(usersByEmail, sub.assignedTo || sub.reviewedBy),
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

function buildPaymentSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedToPayment))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(usersByEmail, sub.assignedToPayment),
            ...(includeDateColumn ? { reportDate: getDateKey(getSubmissionPaymentEntryAt(sub)) || '-' } : {}),
            customerName: sub.customerName || '',
            uploaderName: getUserDisplayNameByEmail(usersByEmail, sub.uploadedBy),
            rsaOfficerName: getUserDisplayNameByEmail(usersByEmail, sub.assignedToRSA),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionPaymentEntryAt(sub)),
        }));
}

function normalizeOutstandingRowForStage(row = {}, stageId = '') {
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
    const lastColLetter = String.fromCharCode(64 + tableHeaders.length);
    worksheet.mergeCells(`A${titleRow.number}:${lastColLetter}${titleRow.number}`);
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
            worksheet.mergeCells(`A${ownerRow.number}:${lastColLetter}${ownerRow.number}`);
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

function normalizeReportScope({ reportDateKey, rangeStartDateKey, rangeEndDateKey }) {
    const singleDate = String(reportDateKey || '').trim();
    const rangeStart = String(rangeStartDateKey || '').trim();
    const rangeEnd = String(rangeEndDateKey || '').trim();
    if (rangeStart && rangeEnd) {
        const start = rangeStart <= rangeEnd ? rangeStart : rangeEnd;
        const end = rangeStart <= rangeEnd ? rangeEnd : rangeStart;
        if (start === REPORT_INCEPTION_START_DATE) {
            return {
                mode: 'from_inception',
                reportDateKey: '',
                rangeStartDateKey: start,
                rangeEndDateKey: end,
                label: `${REPORT_INCEPTION_LABEL} to ${end}`,
                includeDateColumn: true
            };
        }
        return {
            mode: 'date_range',
            reportDateKey: '',
            rangeStartDateKey: start,
            rangeEndDateKey: end,
            label: `${start} to ${end}`,
            includeDateColumn: true
        };
    }
    return {
        mode: 'single_day',
        reportDateKey: singleDate || getPreviousDateKeyInLagos(),
        rangeStartDateKey: '',
        rangeEndDateKey: '',
        label: singleDate || getPreviousDateKeyInLagos(),
        includeDateColumn: false
    };
}

function buildDailyReportDefinition({ submissions = [], users = [], reportDateKey, rangeStartDateKey, rangeEndDateKey }) {
    const usersByEmail = new Map();
    users.forEach((user) => {
        const normalized = normalizeEmail(user?.email);
        if (normalized && !usersByEmail.has(normalized)) usersByEmail.set(normalized, user);
    });

    const scope = normalizeReportScope({ reportDateKey, rangeStartDateKey, rangeEndDateKey });
    const submittedRecords = submissions.filter((sub) => String(sub.status || '').toLowerCase() !== 'draft');
    const reviewerOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return status === 'pending';
    });
    const rsaOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return ['approved', 'processing_to_pfa'].includes(status) && !sub.finalSubmitted && !sub.rsaSubmitted;
    });
    const paymentOutstandingRecords = submittedRecords.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return ['sent_to_pfa', 'rsa_submitted'].includes(status);
    });
    const matchesScope = (value) => ['date_range', 'from_inception'].includes(scope.mode)
        ? isDateKeyWithinRange(value, scope.rangeStartDateKey, scope.rangeEndDateKey)
        : isSameReportDate(value, scope.reportDateKey);
    const uploaderRecords = submittedRecords.filter((sub) => matchesScope(getSubmissionReviewEntryAt(sub)));
    const reviewerRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedTo) && matchesScope(getSubmissionReviewEntryAt(sub)));
    const rsaRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedToRSA) && matchesScope(getSubmissionRsaEntryAt(sub)));
    const paymentRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedToPayment) && matchesScope(getSubmissionPaymentEntryAt(sub)));
    const reviewerAttendedRecords = reviewerRecords.filter((sub) => !!tsToMillis(getSubmissionApprovalEntryAt(sub)));
    const rsaAttendedRecords = rsaRecords.filter((sub) => !!tsToMillis(getSubmissionFinalSubmissionEntryAt(sub)) || String(sub.status || '').toLowerCase() === 'rejected_by_rsa');
    const paymentAttendedRecords = paymentRecords.filter((sub) => !!tsToMillis(getSubmissionPaidEntryAt(sub)) || !!tsToMillis(getSubmissionClearedEntryAt(sub)) || String(sub.status || '').toLowerCase() === 'cleared');
    const dateHeaders = scope.includeDateColumn ? ['Report Date'] : [];
    const dateColumns = scope.includeDateColumn ? [{ width: 14 }] : [];
    return {
        reportDateKey: scope.reportDateKey,
        rangeStartDateKey: scope.rangeStartDateKey,
        rangeEndDateKey: scope.rangeEndDateKey,
        reportLabel: scope.label,
        mode: scope.mode,
        sheets: [
            {
                worksheetName: 'Uploader Report',
                reportTitle: `Uploader Report - ${scope.label}`,
                summaryRows: [
                    ['Total Uploaded', uploaderRecords.length],
                    ['Pending', uploaderRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length]
                ],
                groupRows: buildUploaderSheetRows(uploaderRecords, usersByEmail, scope.includeDateColumn).sort(compareGroupedRows),
                tableHeaders: [...dateHeaders, 'Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Uploaded Time', 'Reject Reason', 'Rejected By', 'Reject Count'],
                columns: [...dateColumns, { width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 24 }, { width: 14 }],
                decimalColumns: scope.includeDateColumn ? [3, 4, 5] : [2, 3, 4],
                integerColumns: scope.includeDateColumn ? [10] : [9]
            },
            {
                worksheetName: 'Reviewer Report',
                reportTitle: `Reviewer Report - ${scope.label}`,
                summaryRows: [
                    ['Total Received', reviewerRecords.length],
                    ['Attended To', reviewerAttendedRecords.length],
                    ['Pending', reviewerRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length],
                    ['Total Outstanding', reviewerOutstandingRecords.length]
                ],
                groupRows: buildReviewerSheetRows(reviewerRecords, usersByEmail, scope.includeDateColumn).sort(compareGroupedRows),
                tableHeaders: [...dateHeaders, 'Customer Name', 'Uploader Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Reject Reason', 'Reject Count'],
                columns: [...dateColumns, { width: 28 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 14 }],
                decimalColumns: scope.includeDateColumn ? [4, 5, 6] : [3, 4, 5],
                integerColumns: scope.includeDateColumn ? [10] : [9]
            },
            {
                worksheetName: 'RSA Report',
                reportTitle: `RSA Report - ${scope.label}`,
                summaryRows: [
                    ['Total Received', rsaRecords.length],
                    ['Attended To', rsaAttendedRecords.length],
                    ['Pending', rsaRecords.filter((sub) => ['approved', 'processing_to_pfa'].includes(String(sub.status || '').toLowerCase()) && !sub.finalSubmitted && !sub.rsaSubmitted).length],
                    ['Total Outstanding', rsaOutstandingRecords.length]
                ],
                groupRows: buildRsaSheetRows(rsaRecords, usersByEmail, scope.includeDateColumn).map(normalizeRsaReportRow).sort(compareGroupedRows),
                tableHeaders: [...dateHeaders, 'Customer Name', 'Uploader Name', 'Reviewer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'RSA Assigned Time'],
                columns: [...dateColumns, { width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
                decimalColumns: scope.includeDateColumn ? [5, 6, 7] : [4, 5, 6],
                integerColumns: []
            },
            {
                worksheetName: 'Payment Report',
                reportTitle: `Payment Report - ${scope.label}`,
                summaryRows: [
                    ['Total Received', paymentRecords.length],
                    ['Attended To', paymentAttendedRecords.length],
                    ['Pending', paymentRecords.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase())).length],
                    ['Total Outstanding', paymentOutstandingRecords.length]
                ],
                groupRows: buildPaymentSheetRows(paymentRecords, usersByEmail, scope.includeDateColumn).sort(compareGroupedRows),
                tableHeaders: [...dateHeaders, 'Customer Name', 'Uploader Name', 'RSA Officer', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time'],
                columns: [...dateColumns, { width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
                decimalColumns: scope.includeDateColumn ? [5, 6, 7] : [4, 5, 6],
                integerColumns: []
            }
        ]
    };
}

function buildOutstandingReportDefinition({ submissions = [], users = [], reportDateKey, rangeStartDateKey, rangeEndDateKey, outstandingDashboard = 'all' }) {
    const usersByEmail = new Map();
    users.forEach((user) => {
        const normalized = normalizeEmail(user?.email);
        if (normalized && !usersByEmail.has(normalized)) usersByEmail.set(normalized, user);
    });

    const scope = normalizeReportScope({ reportDateKey, rangeStartDateKey, rangeEndDateKey });
    const dateKey = scope.label;
    const dashboard = String(outstandingDashboard || 'all').trim().toLowerCase() || 'all';
    const submittedRecords = submissions.filter((sub) => String(sub.status || '').toLowerCase() !== 'draft');
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
            worksheetName: 'Uploader Report',
            reportTitle: `Uploader Outstanding - ${dateKey}`,
            summaryRows: [
                ['Total Outstanding', uploaderOutstandingRecords.length],
                ['Pending', uploaderOutstandingRecords.length]
            ],
            groupRows: buildUploaderSheetRows(uploaderOutstandingRecords, usersByEmail, false).map((row) => normalizeOutstandingRowForStage(row, 'uploader')).sort(compareGroupedRows),
            tableHeaders: ['Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Uploaded Time', 'Reject Reason', 'Rejected By', 'Reject Count'],
            columns: [{ width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 24 }, { width: 14 }],
            decimalColumns: [2, 3, 4],
            integerColumns: [9]
        },
        {
            id: 'reviewer',
            worksheetName: 'Reviewer Report',
            reportTitle: `Reviewer Outstanding - ${dateKey}`,
            summaryRows: [
                ['Total Outstanding', reviewerOutstandingRecords.length],
                ['Pending', reviewerOutstandingRecords.length]
            ],
            groupRows: buildReviewerSheetRows(reviewerOutstandingRecords, usersByEmail, false).map((row) => normalizeOutstandingRowForStage(row, 'reviewer')).sort(compareGroupedRows),
            tableHeaders: ['Customer Name', 'Uploader Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Reject Reason', 'Reject Count'],
            columns: [{ width: 28 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 14 }],
            decimalColumns: [3, 4, 5],
            integerColumns: [9]
        },
        {
            id: 'rsa',
            worksheetName: 'RSA Report',
            reportTitle: `RSA Outstanding - ${dateKey}`,
            summaryRows: [
                ['Total Outstanding', rsaOutstandingRecords.length],
                ['Pending', rsaOutstandingRecords.length]
            ],
            groupRows: buildRsaSheetRows(rsaOutstandingRecords, usersByEmail, false).map((row) => normalizeOutstandingRowForStage(row, 'rsa')).map(normalizeRsaReportRow).sort(compareGroupedRows),
            tableHeaders: ['Customer Name', 'Uploader Name', 'Reviewer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'RSA Assigned Time'],
            columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
            decimalColumns: [4, 5, 6],
            integerColumns: []
        },
        {
            id: 'payment',
            worksheetName: 'Payment Report',
            reportTitle: `Payment Outstanding - ${dateKey}`,
            summaryRows: [
                ['Total Outstanding', paymentOutstandingRecords.length],
                ['Pending', paymentOutstandingRecords.length]
            ],
            groupRows: buildPaymentSheetRows(paymentOutstandingRecords, usersByEmail, false).map((row) => normalizeOutstandingRowForStage(row, 'payment')).sort(compareGroupedRows),
            tableHeaders: ['Customer Name', 'Uploader Name', 'RSA Officer', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Payment Assigned Time'],
            columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
            decimalColumns: [4, 5, 6],
            integerColumns: []
        }
    ];

    const sheets = (dashboard === 'all'
        ? outstandingConfigs
        : outstandingConfigs.filter((config) => config.id === dashboard));

    return {
        reportDateKey: scope.reportDateKey,
        rangeStartDateKey: scope.rangeStartDateKey,
        rangeEndDateKey: scope.rangeEndDateKey,
        reportLabel: `${dashboard === 'all' ? 'All Dashboards' : dashboard.toUpperCase()} Outstanding - ${scope.label}`,
        outstandingDashboard: dashboard,
        sheets
    };
}

async function createDailyReportWorkbookBuffer({ submissions = [], users = [], reportDateKey, rangeStartDateKey, rangeEndDateKey }) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Portal';
    workbook.company = 'CMBank';
    workbook.created = new Date();
    workbook.modified = new Date();

    const report = buildDailyReportDefinition({ submissions, users, reportDateKey, rangeStartDateKey, rangeEndDateKey });
    report.sheets.forEach((sheet) => {
        const worksheet = workbook.addWorksheet(sheet.worksheetName);
        renderGroupedSheet({
            worksheet,
            reportTitle: sheet.reportTitle,
            summaryRows: sheet.summaryRows,
            groupRows: sheet.groupRows,
            tableHeaders: sheet.tableHeaders,
            columns: sheet.columns,
            decimalColumns: sheet.decimalColumns,
            integerColumns: sheet.integerColumns
        });
    });

    return workbook.xlsx.writeBuffer();
}

async function createOutstandingReportWorkbookBuffer({ submissions = [], users = [], reportDateKey, rangeStartDateKey, rangeEndDateKey, outstandingDashboard = 'all' }) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Portal';
    workbook.company = 'CMBank';
    workbook.created = new Date();
    workbook.modified = new Date();

    const report = buildOutstandingReportDefinition({ submissions, users, reportDateKey, rangeStartDateKey, rangeEndDateKey, outstandingDashboard });
    report.sheets.forEach((sheet) => {
        const worksheet = workbook.addWorksheet(sheet.worksheetName);
        renderGroupedSheet({
            worksheet,
            reportTitle: sheet.reportTitle,
            summaryRows: sheet.summaryRows,
            groupRows: sheet.groupRows,
            tableHeaders: sheet.tableHeaders,
            columns: sheet.columns,
            decimalColumns: sheet.decimalColumns,
            integerColumns: sheet.integerColumns
        });
    });

    return workbook.xlsx.writeBuffer();
}

module.exports = {
    createDailyReportWorkbookBuffer,
    createOutstandingReportWorkbookBuffer,
    getPreviousDateKeyInLagos,
    getLagosDateKey
};
