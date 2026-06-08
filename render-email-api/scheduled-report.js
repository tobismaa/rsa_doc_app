const ExcelJS = require('exceljs');

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

function buildUploaderSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records.map((sub) => ({
        owner: getUserDisplayNameByEmail(usersByEmail, sub.uploadedBy),
        ...(includeDateColumn ? { reportDate: getDateKey(sub.uploadedAt) || '-' } : {}),
        customerName: sub.customerName || '',
        rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
        rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
        commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
        status: String(sub.status || '').replace(/_/g, ' '),
        uploadedAt: formatDate(sub.uploadedAt),
        stageTime: formatDate(sub.reviewedAt),
        rejectionReason: getRejectionReason(sub),
        rejectionCount: getRejectionCount(sub)
    }));
}

function buildReviewerSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedTo))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(usersByEmail, sub.assignedTo),
            ...(includeDateColumn ? { reportDate: getDateKey(sub.uploadedAt) || '-' } : {}),
            customerName: sub.customerName || '',
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(sub.uploadedAt),
            stageTime: formatDate(sub.reviewedAt),
            rejectionReason: getRejectionReason(sub),
            rejectionCount: getRejectionCount(sub)
        }));
}

function buildRsaSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedToRSA))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(usersByEmail, sub.assignedToRSA),
            ...(includeDateColumn ? { reportDate: getDateKey(sub.reviewedAt) || '-' } : {}),
            customerName: sub.customerName || '',
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(sub.reviewedAt),
            stageTime: formatDate(sub.finalSubmittedAt || sub.rsaSubmittedAt),
            rejectionReason: String(String(sub.status || '').toLowerCase() === 'rejected_by_rsa' ? getRejectionReason(sub) : ''),
            rejectionCount: Number(String(sub.status || '').toLowerCase() === 'rejected_by_rsa' ? getRejectionCount(sub) : 0)
        }));
}

function buildPaymentSheetRows(records = [], usersByEmail, includeDateColumn = false) {
    return records
        .filter((sub) => normalizeEmail(sub.assignedToPayment))
        .map((sub) => ({
            owner: getUserDisplayNameByEmail(usersByEmail, sub.assignedToPayment),
            ...(includeDateColumn ? { reportDate: getDateKey(sub.paymentAssignedAt || sub.finalSubmittedAt || sub.rsaSubmittedAt) || '-' } : {}),
            customerName: sub.customerName || '',
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(sub.paymentAssignedAt || sub.finalSubmittedAt || sub.rsaSubmittedAt),
            paidAt: formatDate(sub.paidAt),
            clearedAt: formatDate(sub.clearedAt),
            remarks: sub.clearedWithoutAgentCommission ? 'Cleared without agent commission' : String(sub.paymentReconciliationFileName || '').trim()
        }));
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
    const matchesScope = (value) => scope.mode === 'date_range'
        ? isDateKeyWithinRange(value, scope.rangeStartDateKey, scope.rangeEndDateKey)
        : isSameReportDate(value, scope.reportDateKey);
    const uploaderRecords = submittedRecords.filter((sub) => matchesScope(sub.uploadedAt));
    const reviewerRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedTo) && matchesScope(sub.uploadedAt));
    const rsaRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedToRSA) && matchesScope(sub.reviewedAt));
    const paymentRecords = submittedRecords.filter((sub) => normalizeEmail(sub.assignedToPayment) && matchesScope(sub.paymentAssignedAt || sub.finalSubmittedAt || sub.rsaSubmittedAt));
    const dateHeaders = scope.includeDateColumn ? ['Report Date'] : [];
    const dateColumns = scope.includeDateColumn ? [{ width: 14 }] : [];
    const moneyDecimalColumns = scope.includeDateColumn ? [3, 4, 5] : [2, 3, 4];
    const rejectIntegerColumn = scope.includeDateColumn ? [10] : [9];

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
                groupRows: buildUploaderSheetRows(uploaderRecords, usersByEmail, scope.includeDateColumn),
                tableHeaders: [...dateHeaders, 'Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Uploaded Time', 'Reviewer Time', 'Reject Reason', 'Reject Count'],
                columns: [...dateColumns, { width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 22 }, { width: 28 }, { width: 14 }],
                decimalColumns: moneyDecimalColumns,
                integerColumns: rejectIntegerColumn
            },
            {
                worksheetName: 'Reviewer Report',
                reportTitle: `Reviewer Report - ${scope.label}`,
                summaryRows: [
                    ['Total Received', reviewerRecords.length],
                    ['Attending To', reviewerRecords.filter((sub) => normalizeEmail(sub.assignedTo)).length],
                    ['Pending', reviewerRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length]
                ],
                groupRows: buildReviewerSheetRows(reviewerRecords, usersByEmail, scope.includeDateColumn),
                tableHeaders: [...dateHeaders, 'Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Decision Time', 'Reject Reason', 'Reject Count'],
                columns: [...dateColumns, { width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 22 }, { width: 28 }, { width: 14 }],
                decimalColumns: moneyDecimalColumns,
                integerColumns: rejectIntegerColumn
            },
            {
                worksheetName: 'RSA Report',
                reportTitle: `RSA Report - ${scope.label}`,
                summaryRows: [
                    ['Total Received', rsaRecords.length],
                    ['Attending To', rsaRecords.filter((sub) => ['approved', 'processing_to_pfa'].includes(String(sub.status || '').toLowerCase()) && !sub.finalSubmitted && !sub.rsaSubmitted).length],
                    ['Pending', rsaRecords.filter((sub) => ['approved', 'processing_to_pfa'].includes(String(sub.status || '').toLowerCase()) && !sub.finalSubmitted && !sub.rsaSubmitted).length]
                ],
                groupRows: buildRsaSheetRows(rsaRecords, usersByEmail, scope.includeDateColumn),
                tableHeaders: [...dateHeaders, 'Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'RSA Assigned Time', 'RSA Done Time', 'RSA Reject Reason', 'RSA Reject Count'],
                columns: [...dateColumns, { width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 22 }, { width: 28 }, { width: 14 }],
                decimalColumns: moneyDecimalColumns,
                integerColumns: rejectIntegerColumn
            },
            {
                worksheetName: 'Payment Report',
                reportTitle: `Payment Report - ${scope.label}`,
                summaryRows: [
                    ['Total Received', paymentRecords.length],
                    ['Attending To', paymentRecords.filter((sub) => normalizeEmail(sub.assignedToPayment)).length],
                    ['Pending', paymentRecords.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase())).length]
                ],
                groupRows: buildPaymentSheetRows(paymentRecords, usersByEmail, scope.includeDateColumn),
                tableHeaders: [...dateHeaders, 'Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Paid Time', 'Cleared Time', 'Remarks'],
                columns: [...dateColumns, { width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 22 }, { width: 22 }, { width: 28 }],
                decimalColumns: moneyDecimalColumns,
                integerColumns: []
            }
        ]
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

module.exports = {
    createDailyReportWorkbookBuffer,
    getPreviousDateKeyInLagos,
    getLagosDateKey
};
