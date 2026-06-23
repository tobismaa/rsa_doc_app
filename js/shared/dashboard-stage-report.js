import { formatAppDateTime } from './app-time.js';
import {
    getTimestampMillis,
    getSubmissionReviewEntryAt,
    getSubmissionApprovalEntryAt,
    getSubmissionRsaEntryAt,
    getSubmissionFinalSubmissionEntryAt,
    getSubmissionPaymentEntryAt,
    getSubmissionPaidEntryAt,
    getSubmissionClearedEntryAt
} from './submission-stage.js?v=20260609a';

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

function getDateKey(value) {
    const ms = getTimestampMillis(value);
    if (!ms) return '';
    const date = new Date(ms);
    const year = date.getFullYear();
    const month = `${date.getMonth() + 1}`.padStart(2, '0');
    const day = `${date.getDate()}`.padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function isDateWithinRange(value, startDateKey = '', endDateKey = '') {
    const dateKey = getDateKey(value);
    if (!dateKey) return false;
    const start = String(startDateKey || '').trim();
    const end = String(endDateKey || '').trim();
    if (start && dateKey < start) return false;
    if (end && dateKey > end) return false;
    return true;
}

function formatDate(value) {
    return formatAppDateTime(value, '-');
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

function getRejectionReason(sub = {}) {
    return String(sub?.latestRejectionReason || sub?.previousRejectionReason || sub?.comment || '').trim();
}

function getRejectionCount(sub = {}) {
    const history = Array.isArray(sub?.rejectionHistory) ? sub.rejectionHistory.filter((entry) => {
        if (typeof entry === 'string') return String(entry).trim();
        return String(entry?.reason || '').trim();
    }) : [];
    if (history.length) return history.length;
    return getRejectionReason(sub) ? 1 : 0;
}

function getRejectionOfficerName(sub = {}, resolveName = () => 'Unassigned') {
    if (getRejectionCount(sub) <= 0 && !getRejectionReason(sub)) return '';
    const rejectedBy = String(sub?.latestRejectedBy || sub?.previousRejectedBy || '').trim();
    if (rejectedBy) return resolveName(rejectedBy);
    const stage = String(sub?.latestRejectedStage || '').trim().toLowerCase();
    if (stage === 'rsa') return resolveName(sub?.assignedToRSA || '');
    if (stage === 'payment') return resolveName(sub?.assignedToPayment || '');
    return resolveName(sub?.reviewedBy || sub?.assignedTo || '');
}

function getCustomerDetailsValue(sub = {}, keys = []) {
    const details = sub?.customerDetails && typeof sub.customerDetails === 'object' ? sub.customerDetails : {};
    for (const key of keys) {
        const value = details?.[key] ?? sub?.[key];
        if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
    }
    return '';
}

function compareGroupedRows(a = {}, b = {}) {
    const ownerA = String(a.owner || '').toLowerCase();
    const ownerB = String(b.owner || '').toLowerCase();
    if (ownerA !== ownerB) return ownerA.localeCompare(ownerB);

    const assignedA = String(a.assignedAt || a.uploadedAt || '').toLowerCase();
    const assignedB = String(b.assignedAt || b.uploadedAt || '').toLowerCase();
    if (assignedA !== assignedB) return assignedA.localeCompare(assignedB);

    const stageA = String(a.stageTime || '').toLowerCase();
    const stageB = String(b.stageTime || '').toLowerCase();
    if (stageA !== stageB) return stageA.localeCompare(stageB);

    return String(a.customerName || '').toLowerCase().localeCompare(String(b.customerName || '').toLowerCase());
}

function normalizeRsaReportRow(row = {}) {
    return {
        owner: row.owner || '',
        customerName: row.customerName || '',
        uploaderName: row.uploaderName || '',
        reviewerName: row.reviewerName || '',
        accountNumber: row.accountNumber || '',
        tenor: row.tenor || '',
        houseType: row.houseType || '',
        houseNumber: row.houseNumber || '',
        rsaBalance: row.rsaBalance ?? '',
        rsa25: row.rsa25 ?? '',
        commission: row.commission ?? '',
        status: row.status || '',
        assignedAt: row.assignedAt || '-'
    };
}

function buildUploaderSheetRows(records = [], resolveName = () => 'Unassigned') {
    return records.map((sub) => ({
        owner: resolveName(sub.uploadedBy),
        customerName: sub.customerName || '',
        rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
        rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
        commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
        status: String(sub.status || '').replace(/_/g, ' '),
        uploadedAt: formatDate(getSubmissionReviewEntryAt(sub)),
        rejectionReason: getRejectionReason(sub),
        rejectionOfficer: getRejectionOfficerName(sub, resolveName),
        rejectionCount: getRejectionCount(sub)
    }));
}

function buildReviewerSheetRows(records = [], resolveName = () => 'Unassigned') {
    return records
        .filter((sub) => String(sub?.assignedTo || '').trim())
        .map((sub) => ({
            owner: resolveName(sub.assignedTo),
            customerName: sub.customerName || '',
            uploaderName: resolveName(sub.uploadedBy),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionReviewEntryAt(sub)),
            rejectionReason: getRejectionReason(sub),
            rejectionCount: getRejectionCount(sub)
        }));
}

function buildRsaSheetRows(records = [], resolveName = () => 'Unassigned') {
    return records
        .filter((sub) => String(sub?.assignedToRSA || '').trim())
        .map((sub) => ({
            owner: resolveName(sub.assignedToRSA),
            customerName: sub.customerName || '',
            uploaderName: resolveName(sub.uploadedBy),
            reviewerName: resolveName(sub.assignedTo || sub.reviewedBy),
            accountNumber: getCustomerDetailsValue(sub, ['accountNo', 'accountNumber']),
            tenor: getCustomerDetailsValue(sub, ['tenor']),
            houseType: getCustomerDetailsValue(sub, ['propertyType', 'houseType']),
            houseNumber: getCustomerDetailsValue(sub, ['houseNumber']),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionRsaEntryAt(sub)),
            stageTime: formatDate(getSubmissionFinalSubmissionEntryAt(sub))
        }));
}

function buildPaymentSheetRows(records = [], resolveName = () => 'Unassigned') {
    return records
        .filter((sub) => String(sub?.assignedToPayment || '').trim())
        .map((sub) => ({
            owner: resolveName(sub.assignedToPayment),
            customerName: sub.customerName || '',
            uploaderName: resolveName(sub.uploadedBy),
            rsaOfficerName: resolveName(sub.assignedToRSA),
            rsaBalance: formatMoneyForSheet(getSubmissionRsaBalance(sub)),
            rsa25: formatMoneyForSheet(getSubmissionTwentyFivePercent(sub)),
            commission: formatMoneyForSheet(getSubmissionCommissionOnePercent(sub)),
            status: String(sub.status || '').replace(/_/g, ' '),
            assignedAt: formatDate(getSubmissionPaymentEntryAt(sub))
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

function buildPreviewValues(sheetId, row) {
    if (sheetId === 'reviewer') {
        return [row.customerName, row.uploaderName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt, String(row.rejectionCount || 0)];
    }
    if (sheetId === 'rsa') {
        return [row.customerName, row.uploaderName, row.reviewerName, row.accountNumber, row.tenor, row.houseType, row.houseNumber, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt];
    }
    if (sheetId === 'payment') {
        return [row.customerName, row.uploaderName, row.rsaOfficerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt];
    }
    return [row.customerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.uploadedAt, row.rejectionOfficer, String(row.rejectionCount || 0)];
}

function buildDashboardTableValues(sheetId, row) {
    if (sheetId === 'reviewer') {
        return [row.owner, row.customerName, row.uploaderName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt, row.rejectionReason, String(row.rejectionCount || 0)];
    }
    if (sheetId === 'rsa') {
        return [row.owner, row.customerName, row.uploaderName, row.reviewerName, row.accountNumber, row.tenor, row.houseType, row.houseNumber, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt];
    }
    if (sheetId === 'payment') {
        return [row.owner, row.customerName, row.uploaderName, row.rsaOfficerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.assignedAt];
    }
    return [row.owner, row.customerName, formatMoneyPreview(row.rsaBalance), formatMoneyPreview(row.rsa25), formatMoneyPreview(row.commission), row.status, row.uploadedAt, row.rejectionReason, row.rejectionOfficer, String(row.rejectionCount || 0)];
}

function getStageSheetConfig(stageId = '') {
    const stage = String(stageId || '').trim().toLowerCase();
    if (stage === 'reviewer') {
        return {
            id: 'reviewer',
            title: 'Reviewer Report',
            excelHeaders: ['Customer Name', 'Uploader Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Assigned Time', 'Reject Reason', 'Reject Count'],
            previewHeaders: ['Customer', 'Uploader', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned', 'Reject Count'],
            columns: [{ width: 28 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 14 }],
            decimalColumns: [3, 4, 5],
            integerColumns: [9]
        };
    }
    if (stage === 'rsa') {
        return {
            id: 'rsa',
            title: 'RSA Report',
            excelHeaders: ['Customer Name', 'Uploader Name', 'Reviewer Name', 'Account Number', 'Tenor', 'House Type', 'House Number', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'RSA Assigned Time'],
            previewHeaders: ['Customer', 'Uploader', 'Reviewer', 'Account No', 'Tenor', 'House Type', 'House No', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned'],
            columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 18 }, { width: 12 }, { width: 28 }, { width: 18 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
            decimalColumns: [8, 9, 10],
            integerColumns: []
        };
    }
    if (stage === 'payment') {
        return {
            id: 'payment',
            title: 'Payment Report',
            excelHeaders: ['Customer Name', 'Uploader Name', 'RSA Officer', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Payment Assigned Time'],
            previewHeaders: ['Customer', 'Uploader', 'RSA Officer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Assigned'],
            columns: [{ width: 28 }, { width: 24 }, { width: 24 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }],
            decimalColumns: [4, 5, 6],
            integerColumns: []
        };
    }
    return {
        id: 'uploader',
        title: 'Uploader Report',
        excelHeaders: ['Customer Name', 'RSA Balance', '25% RSA Balance', '1% Commission', 'Status', 'Uploaded Time', 'Reject Reason', 'Rejected By', 'Reject Count'],
        previewHeaders: ['Customer', 'RSA Bal', '25%', '1% Comm', 'Status', 'Uploaded', 'Rejected By', 'Reject Count'],
        columns: [{ width: 28 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 28 }, { width: 24 }, { width: 14 }],
        decimalColumns: [2, 3, 4],
        integerColumns: [9]
    };
}

export function buildDashboardStageReport({ stageId = '', records = [], rangeStart = '', rangeEnd = '', resolveName = (email) => String(email || '').trim() || 'Unassigned' } = {}) {
    const stage = String(stageId || '').trim().toLowerCase() || 'uploader';
    const dateRangeLabel = rangeStart && rangeEnd ? `${rangeStart} to ${rangeEnd}` : rangeStart || rangeEnd || '-';
    const sourceRecords = records.filter((sub) => {
        if (stage === 'reviewer') return isDateWithinRange(getSubmissionReviewEntryAt(sub), rangeStart, rangeEnd);
        if (stage === 'rsa') return isDateWithinRange(getSubmissionRsaEntryAt(sub), rangeStart, rangeEnd);
        if (stage === 'payment') return isDateWithinRange(getSubmissionPaymentEntryAt(sub), rangeStart, rangeEnd);
        return isDateWithinRange(getSubmissionReviewEntryAt(sub), rangeStart, rangeEnd);
    });

    let rows = [];
    let summaryRows = [];
    const config = getStageSheetConfig(stage);

    if (stage === 'reviewer') {
        rows = buildReviewerSheetRows(sourceRecords, resolveName).sort(compareGroupedRows);
        const attended = sourceRecords.filter((sub) => !!getTimestampMillis(getSubmissionApprovalEntryAt(sub))).length;
        const pending = sourceRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length;
        summaryRows = [
            ['Total Received', sourceRecords.length],
            ['Attended To', attended],
            ['Pending', pending],
            ['Date Range', dateRangeLabel]
        ];
    } else if (stage === 'rsa') {
        rows = buildRsaSheetRows(sourceRecords, resolveName).map(normalizeRsaReportRow).sort(compareGroupedRows);
        const attended = sourceRecords.filter((sub) => !!getTimestampMillis(getSubmissionFinalSubmissionEntryAt(sub)) || String(sub.status || '').toLowerCase() === 'rejected_by_rsa').length;
        const pending = sourceRecords.filter((sub) => {
            const status = String(sub.status || '').toLowerCase();
            return ['approved', 'processing_to_pfa'].includes(status) && !sub.finalSubmitted && !sub.rsaSubmitted;
        }).length;
        summaryRows = [
            ['Total Received', sourceRecords.length],
            ['Attended To', attended],
            ['Pending', pending],
            ['Date Range', dateRangeLabel]
        ];
    } else if (stage === 'payment') {
        rows = buildPaymentSheetRows(sourceRecords, resolveName).sort(compareGroupedRows);
        const attended = sourceRecords.filter((sub) => !!getTimestampMillis(getSubmissionPaidEntryAt(sub)) || !!getTimestampMillis(getSubmissionClearedEntryAt(sub)) || String(sub.status || '').toLowerCase() === 'cleared').length;
        const pending = sourceRecords.filter((sub) => ['sent_to_pfa', 'rsa_submitted'].includes(String(sub.status || '').toLowerCase())).length;
        summaryRows = [
            ['Total Received', sourceRecords.length],
            ['Attended To', attended],
            ['Pending', pending],
            ['Date Range', dateRangeLabel]
        ];
    } else {
        rows = buildUploaderSheetRows(sourceRecords, resolveName).sort(compareGroupedRows);
        const pending = sourceRecords.filter((sub) => String(sub.status || '').toLowerCase() === 'pending').length;
        summaryRows = [
            ['Total Uploaded', sourceRecords.length],
            ['Pending', pending],
            ['Rejected', sourceRecords.filter((sub) => getRejectionCount(sub) > 0).length],
            ['Date Range', dateRangeLabel]
        ];
    }

    return {
        stageId: stage,
        title: `${config.title} - ${dateRangeLabel}`,
        metaText: `${config.title} generated for ${dateRangeLabel}.`,
        rangeStart,
        rangeEnd,
        filePrefix: `${stage}_report`,
        excelHeaders: config.excelHeaders,
        previewHeaders: config.previewHeaders,
        columns: config.columns,
        decimalColumns: config.decimalColumns,
        integerColumns: config.integerColumns,
        summaryRows,
        rows,
        previewRows: rows.map((row) => ({
            owner: row.owner,
            previewValues: buildPreviewValues(stage, row)
        }))
    };
}

export function renderDashboardStageReport(report, refs = {}) {
    if (!report) return;
    const { metaEl, summaryBodyEl, detailsBodyEl } = refs;
    if (metaEl) metaEl.textContent = report.metaText || 'Report preview ready.';
    if (summaryBodyEl) {
        summaryBodyEl.innerHTML = report.summaryRows.length
            ? report.summaryRows.map(([label, value]) => `<tr><td>${String(label || '')}</td><td>${String(value ?? '')}</td></tr>`).join('')
            : '<tr><td colspan="2" class="no-data">No summary available.</td></tr>';
    }
    if (detailsBodyEl) {
        const displayRows = report.rows.map((row) => buildDashboardTableValues(report.stageId, row));
        detailsBodyEl.innerHTML = displayRows.length
            ? displayRows.map((values) => `<tr>${values.map((value) => `<td>${String(value ?? '-')}</td>`).join('')}</tr>`).join('')
            : `<tr><td colspan="${displayRows[0]?.length || (report.stageId === 'rsa' ? 13 : (report.stageId === 'uploader' || report.stageId === 'reviewer' ? 10 : 9))}" class="no-data">No records found for the selected date range.</td></tr>`;
    }
}

export async function exportDashboardStageReportExcel(report, creatorName = 'CMBank RSA Dashboard') {
    if (!report || !window.ExcelJS) {
        throw new Error('Excel library is not available right now.');
    }

    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = creatorName;
    workbook.created = new Date();

    renderGroupedSheet({
        worksheet: workbook.addWorksheet(`${report.stageId} Report`),
        reportTitle: report.title,
        summaryRows: report.summaryRows,
        groupRows: report.rows,
        tableHeaders: report.excelHeaders,
        columns: report.columns,
        decimalColumns: report.decimalColumns,
        integerColumns: report.integerColumns
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const suffix = `${report.rangeStart || 'start'}_${report.rangeEnd || 'end'}`.replace(/[^\w-]/g, '_');
    link.download = `${report.filePrefix}_${suffix}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
}
