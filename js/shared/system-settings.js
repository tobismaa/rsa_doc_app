import { doc, getDoc } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

const CACHE_TTL_MS = 30 * 1000;
let settingsCache = null;
let settingsFetchedAt = 0;
const DEFAULT_AGENT_BANK_NAMES = [
  'Access Bank',
  'ALAT by Wema',
  'Coronation Merchant Bank',
  'Citibank',
  'Ecobank',
  'FairMoney MFB',
  'Fidelity Bank',
  'First Bank of Nigeria',
  'First City Monument Bank (FCMB)',
  'FSDH Merchant Bank',
  'Globus Bank',
  'Guaranty Trust Bank (GTBank)',
  'Heritage Bank',
  'Jaiz Bank',
  'Keystone Bank',
  'Kuda MFB',
  'Lotus Bank',
  'Moniepoint MFB',
  'Optimus Bank',
  'Opay',
  'PalmPay',
  'Parallex Bank',
  'Polaris Bank',
  'PremiumTrust Bank',
  'Providus Bank',
  'Rand Merchant Bank',
  'Rubies MFB',
  'Safe Haven MFB',
  'Signature Bank',
  'Stanbic IBTC Bank',
  'Standard Chartered Bank',
  'Sterling Bank',
  'SunTrust Bank',
  'TajBank',
  'Titan Trust Bank',
  'Union Bank of Nigeria',
  'United Bank for Africa (UBA)',
  'Unity Bank',
  'VFD Microfinance Bank',
  'Wema Bank',
  'Woven Finance',
  'Zenith Bank'
];

const DEFAULT_PFA_NAMES = [
  'Access ARM Pensions Limited',
  'ACCESS ARM',
  'Cardinal Stone Pensions Limited',
  'Citizens Pensions Limited',
  'Crusader Sterling Pensions Limited',
  'CRUSADER',
  'FCMB Pensions Limited',
  'FCMB Pensions',
  'Fidelity Pension Managers Limited',
  'Guaranty Trust Pension Managers Limited',
  'GTCO Pension Managers Limited',
  'Leadway Pensure PFA Limited',
  'LEADWAY PENSURE PFA',
  'NLPC Pension Fund Administrators Limited',
  'NLPC Pensions',
  'Nigerian University Pension Management Company (NUPEMCO)',
  'NUPEMCO Pensions',
  'Norrenberger Pensions Limited',
  'NPF Pension Managers Limited',
  'NPF Pensions',
  'OAK Pensions Limited',
  'OAK Pensions',
  'Parthian Pensions Limited',
  'Pensions Alliance Limited',
  'PAL Pensions',
  'Premium Pension Limited',
  'Stanbic IBTC Pension Managers Limited',
  'Stanbic IBTC Pensions',
  'Tangerine APT Pensions Limited',
  'Trustfund Pensions Limited',
  'TRUSTFUND',
  'Veritas Glanvills Pensions Limited'
];

const DEFAULT_DOCUMENT_REQUIREMENTS = [
  { id: 'birth_certificate', name: 'Birth Certificate / Age Declaration', icon: 'fa-id-card', required: true, active: true },
  { id: 'nin', name: 'National Identification Number (NIN)', icon: 'fa-id-card', required: true, active: true },
  { id: 'bvn', name: 'BVN', icon: 'fa-id-badge', required: true, active: true },
  { id: 'pay_slips', name: '3 Months Pay Slip', icon: 'fa-file-invoice', required: false, active: true },
  { id: 'offer_letter', name: 'Offer of Employment Letter', icon: 'fa-file-signature', required: true, active: true },
  { id: 'intro_letter', name: 'Introduction Letter', icon: 'fa-file-signature', required: true, active: true },
  { id: 'request_letter', name: 'Request Letter', icon: 'fa-file-signature', required: true, active: true },
  { id: 'rsa_statement', name: 'RSA Statement', icon: 'fa-file-invoice', required: true, active: true },
  { id: 'pfa_form', name: 'PFA Application Form', icon: 'fa-file-pdf', required: true, active: true },
  { id: 'consent_letter', name: 'Consent Letter', icon: 'fa-file-signature', required: true, active: true },
  { id: 'indemnity_form', name: 'Indemnity Form', icon: 'fa-file-signature', required: true, active: true },
  { id: 'mortgage_loan_application', name: 'Mortgage Loan Application Form', icon: 'fa-file-pdf', required: true, active: true },
  { id: 'allocation_last_page', name: 'Allocation Last Page', icon: 'fa-file-pdf', required: true, active: true },
  { id: 'offer_letter_last_page', name: 'Offer Letter Last Page', icon: 'fa-file-pdf', required: true, active: true },
  { id: 'pmi_soa', name: 'PMI SOA', icon: 'fa-file-pdf', required: true, active: true },
  { id: 'benefit_application_form', name: 'Benefit Application Form', icon: 'fa-file-alt', required: false, active: true },
  { id: 'data_recapture', name: 'Data Recapture', icon: 'fa-file-alt', required: false, active: true },
  { id: 'credit_life', name: 'Credit Life', icon: 'fa-file-medical', required: false, active: true }
];

const DEFAULT_PROPERTY_RULES = [
  { name: '1 BEDROOM 8 IN 1 FLAT', value: 6000000, fee: 2000, min: 4000000, max: 6499999 },
  { name: '1 BEDROOM 4 IN 1 BUNGALOW', value: 6500000, fee: 2000, min: 6500000, max: 6999999 },
  { name: '1 BEDROOM 2 IN 1 BUNGALOW', value: 7000000, fee: 2000, min: 7000000, max: 9999999 },
  { name: '2 BEDROOM SEMI DETACHED BUNGALOW', value: 13000000, fee: 3000, min: 10000000, max: 14999999 },
  { name: '3 BEDROOM SEMI DETACHED BUNGALOW', value: 15000000, fee: 4000, min: 15000000, max: 19999999 },
  { name: '4 BEDROOM DETACHED BUNGALOW', value: 24000000, fee: 5000, min: 20000000, max: 34999999 },
  { name: '4 BEDROOM DETACHED LUXURY BUNGALOW', value: 35000000, fee: 5000, min: 35000000, max: 59999999 },
  { name: '4 BEDROOM TERRACE DUPLEX', value: 60000000, fee: 10000, min: 60000000, max: 99999999 },
  { name: '5 BEDROOM TERRACE DUPLEX', value: 100000000, fee: 20000, min: 100000000, max: 149999999 },
  { name: '6 BEDROOM TERRACE DUPLEX', value: 150000000, fee: 30000, min: 150000000, max: 199999999 },
  { name: '7 BEDROOM TERRACE DUPLEX', value: 200000000, fee: 40000, min: 200000000, max: 249999999 },
  { name: '8 BEDROOM TERRACE DUPLEX', value: 250000000, fee: 50000, min: 250000000, max: 299999999 }
];

const DEFAULT_HOUSE_NUMBER_RULES = {
  '1 BEDROOM 8 IN 1 FLAT': { mode: 'alpha_suffix', prefix: 'C', startNumber: 12, startLetter: 'M' },
  '1 BEDROOM 2 IN 1 BUNGALOW': { mode: 'alpha_suffix', prefix: 'J', startNumber: 55, startLetter: 'A' },
  '1 BEDROOM 4 IN 1 BUNGALOW': { mode: 'alpha_suffix', prefix: 'J', startNumber: 50, startLetter: 'M' },
  '2 BEDROOM SEMI DETACHED BUNGALOW': { mode: 'alpha_suffix', prefix: 'X', startNumber: 60, startLetter: 'A' },
  '3 BEDROOM SEMI DETACHED BUNGALOW': { mode: 'block_100', startPrefix: 'M', startNumber: 60 },
  '4 BEDROOM DETACHED BUNGALOW': { mode: 'block_100', startPrefix: 'N', startNumber: 71 },
  '4 BEDROOM DETACHED LUXURY BUNGALOW': { mode: 'block_100', startPrefix: 'P', startNumber: 26 },
  '4 BEDROOM TERRACE DUPLEX': { mode: 'house_infinite', startNumber: 20 },
  '5 BEDROOM TERRACE DUPLEX': { mode: 'house_block_100', startPrefix: 'B', startNumber: 6 },
  '6 BEDROOM TERRACE DUPLEX': { mode: 'house_block_100', startPrefix: 'A', startNumber: 12 }
};

const DEFAULT_BULK_IMPORT_COLUMNS = [
  'Customer Name',
  'Date of Birth',
  'Account No',
  'Email',
  'Phone',
  'Agent Name',
  'NIN',
  'Address',
  'Employer',
  'Originating TP',
  'Mortgage Form Date',
  'PFA',
  'PEN No',
  'RSA Statement Date',
  'RSA Balance'
];

function parseBoolean(value, fallback = false) {
  return typeof value === 'boolean' ? value : fallback;
}

function parseNumber(value, fallback = 0) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function parseText(value, fallback = '') {
  const text = String(value || '').trim();
  return text || fallback;
}

function parseHexColor(value, fallback = '') {
  const text = String(value || '').trim();
  return /^#[0-9a-f]{6}$/i.test(text) ? text : fallback;
}

function parseDurationSetting(value, fallback = '11m') {
  const text = String(value || '').trim().toLowerCase();
  if (/^\d+\s*[smh]$/.test(text)) return text.replace(/\s+/g, '');
  if (/^\d+(\.\d+)?$/.test(text)) return `${text}m`;
  return fallback;
}

function parseStringArray(value, fallback = []) {
  const source = Array.isArray(value) ? value : fallback;
  const seen = new Set();
  const items = [];
  source.forEach((entry) => {
    const text = parseText(entry, '');
    if (!text) return;
    const key = text.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    items.push(text);
  });
  return items;
}

function normalizeAnnouncementTarget(value = '') {
  const text = String(value || '').trim().toLowerCase();
  if (['reports_monitoring', 'reports-monitoring', 'reporting_monitoring', 'reporting-monitoring'].includes(text)) return 'audit';
  return text;
}

function parseAnnouncementTargets(value, fallback = []) {
  return Array.from(new Set(
    parseStringArray(value, fallback)
      .map(normalizeAnnouncementTarget)
      .filter(Boolean)
  ));
}

function parseObject(value, fallback = {}) {
  return value && typeof value === 'object' && !Array.isArray(value) ? value : fallback;
}

function normalizeHouseNumberRules(value, fallback = {}) {
  const base = parseObject(fallback, {});
  const source = parseObject(value, {});
  const merged = Object.keys(source).length ? { ...base, ...source } : base;
  return Object.fromEntries(
    Object.entries(merged)
      .map(([propertyType, rule]) => {
        const name = parseText(propertyType, '');
        if (!name || !rule || typeof rule !== 'object' || Array.isArray(rule)) return null;
        return [name, {
          mode: parseText(rule.mode, 'alpha_suffix'),
          prefix: parseText(rule.prefix, ''),
          startNumber: parseNumber(rule.startNumber, 0),
          startLetter: parseText(rule.startLetter, ''),
          startPrefix: parseText(rule.startPrefix, '')
        }];
      })
      .filter(Boolean)
  );
}

function normalizeDocumentRequirementRoles(value, fallback = {}) {
  const base = parseObject(fallback, {});
  const source = parseObject(value, {});
  const merged = { ...base, ...source };
  return {
    uploader_level_1: parseBoolean(merged.uploader_level_1, parseBoolean(merged.uploader, true)),
    uploader_level_2: parseBoolean(merged.uploader_level_2, parseBoolean(merged.uploader, false)),
    reviewer: parseBoolean(merged.reviewer, true),
    rsa_level_1: parseBoolean(merged.rsa_level_1, parseBoolean(merged.rsa, false)),
    rsa_level_2: parseBoolean(merged.rsa_level_2, parseBoolean(merged.rsa, false)),
    admin: parseBoolean(merged.admin, true),
    super_admin: parseBoolean(merged.super_admin, true),
    payment: parseBoolean(merged.payment, true)
  };
}

export function normalizeAgentBankOptions(value, fallback = []) {
  const source = Array.isArray(value) ? value : fallback;
  const seen = new Set();
  const options = [];

  source.forEach((entry) => {
    const bankName = typeof entry === 'string'
      ? parseText(entry, '')
      : parseText(entry?.name, '');
    if (!bankName) return;

    const dedupeKey = bankName.toLowerCase();
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    options.push({
      name: bankName,
      active: typeof entry === 'object' && entry !== null
        ? parseBoolean(entry.active, true)
        : true
    });
  });

  return options;
}

export function normalizeDocumentRequirements(value, fallback = []) {
  const source = Array.isArray(value) ? value : fallback;
  const seen = new Set();
  const items = [];

  source.forEach((entry) => {
    const id = parseText(entry?.id, '');
    const name = parseText(entry?.name, '');
    if (!id || !name) return;
    if (seen.has(id)) return;
    seen.add(id);
    items.push({
      id,
      name,
      icon: parseText(entry?.icon, 'fa-file-alt'),
      required: parseBoolean(entry?.required, false),
      active: parseBoolean(entry?.active, true)
    });
  });

  return items;
}

export function normalizePropertyRules(value, fallback = []) {
  const source = Array.isArray(value) ? value : fallback;
  return source
    .map((entry) => ({
      name: parseText(entry?.name, ''),
      value: parseNumber(entry?.value, 0),
      fee: parseNumber(entry?.fee, 0),
      min: parseNumber(entry?.min, 0),
      max: parseNumber(entry?.max, 0)
    }))
    .filter((entry) => entry.name && entry.max >= entry.min);
}

export function getDefaultSystemSettings() {
  const agentBankOptions = normalizeAgentBankOptions(DEFAULT_AGENT_BANK_NAMES);
  const pfaOptions = parseStringArray(DEFAULT_PFA_NAMES);
  const documentRequirements = normalizeDocumentRequirements(DEFAULT_DOCUMENT_REQUIREMENTS);
  const propertyRules = normalizePropertyRules(DEFAULT_PROPERTY_RULES);
  return {
    maintenanceMode: false,
    maintenanceMessage: 'System is currently under maintenance. Please try again later.',
    commissionRate: 0.01,
    commissionRateEffectiveFrom: '2026-05-07T00:00:00+01:00',
    maxImageUploadMb: 1,
    maxPdfUploadMb: 1.5,
    reviewerRoundRobinEnabled: true,
    rsaRoundRobinEnabled: true,
    paymentRoundRobinEnabled: true,
    agentEditSyncEnabled: true,
    notificationsEmailEnabled: true,
    notificationsPushEnabled: true,
    scheduledReportEmail: {
      enabled: false,
      subject: 'Daily RSA Report',
      body: 'Hello,\n\nPlease find the attached daily RSA report.\n\nRegards,\nCMBank RSA Portal',
      recipients: [],
      sendTime: '08:00',
      reportDateMode: 'previous_day'
    },
    agentBankOptions,
    pfaOptions,
    pfaAddresses: {},
    documentRequirements,
    documentRequirementRoles: normalizeDocumentRequirementRoles(),
    rolePermissions: {
      uploaderCanUpload: true,
      reviewerReviewAccessEnabled: true,
      reviewerCanApprove: true,
      reviewerCanReject: true,
      rsaCanApprove: true,
      rsaCanReject: true,
      superAdminCanEditAgentRecords: true,
      superAdminCanClearCache: true
    },
    routingPolicies: {
      defaultRouteMode: 'normal',
      fallbackAssignmentMode: 'round_robin'
    },
    workflowLabels: {
      pending: 'Pending',
      processing_to_pfa: 'Processing to PFA',
      approved: 'Approved',
      rejected: 'Rejected',
      rejected_by_rsa: 'Rejected by RSA',
      sent_to_pfa: 'Sent to PFA',
      paid: 'Paid',
      cleared: 'Cleared'
    },
    rejectionRules: {
      reviewerRequired: true,
      rsaRequired: true,
      minLength: 10,
      reviewerCannedReasons: [],
      rsaCannedReasons: []
    },
    agentRegistrationRules: {
      approvalRequired: true
    },
    bulkImportRules: {
      requiredColumns: parseStringArray(DEFAULT_BULK_IMPORT_COLUMNS)
    },
    uploadControls: {
      simulateStorageCapFullFailure: false
    },
    dashboardAnnouncement: {
      enabled: false,
      message: '',
      tone: 'info',
      speed: 30,
      textColor: '',
      fontSize: 15,
      fontStyle: 'bold',
      fontFamily: 'system',
      targetDashboards: []
    },
    globalReadOnlyMode: false,
    globalReadOnlyMessage: 'Read-only mode is active. You can view records, but changes are temporarily disabled.',
    securityControls: {
      sessionTimeoutMinutes: 60,
      forceLogoutCountdown: '11m',
      forceLogoutToken: ''
    },
    notificationTemplates: {},
    auditControls: {
      retentionDays: 30
    },
    propertyRules,
    houseNumberRules: normalizeHouseNumberRules(DEFAULT_HOUSE_NUMBER_RULES, {})
  };
}

function normalizeSystemSettings(data = {}) {
  const defaults = getDefaultSystemSettings();
  return {
    maintenanceMode: parseBoolean(data.maintenanceMode, defaults.maintenanceMode),
    maintenanceMessage: parseText(data.maintenanceMessage, defaults.maintenanceMessage),
    commissionRate: parseNumber(data.commissionRate, defaults.commissionRate),
    commissionRateEffectiveFrom: parseText(data.commissionRateEffectiveFrom, defaults.commissionRateEffectiveFrom),
    maxImageUploadMb: Math.max(0.1, parseNumber(data.maxImageUploadMb, defaults.maxImageUploadMb)),
    maxPdfUploadMb: Math.max(0.1, parseNumber(data.maxPdfUploadMb, defaults.maxPdfUploadMb)),
    reviewerRoundRobinEnabled: parseBoolean(data.reviewerRoundRobinEnabled, defaults.reviewerRoundRobinEnabled),
    rsaRoundRobinEnabled: parseBoolean(data.rsaRoundRobinEnabled, defaults.rsaRoundRobinEnabled),
    paymentRoundRobinEnabled: parseBoolean(data.paymentRoundRobinEnabled, defaults.paymentRoundRobinEnabled),
    agentEditSyncEnabled: parseBoolean(data.agentEditSyncEnabled, defaults.agentEditSyncEnabled),
    notificationsEmailEnabled: parseBoolean(data.notificationsEmailEnabled, defaults.notificationsEmailEnabled),
    notificationsPushEnabled: parseBoolean(data.notificationsPushEnabled, defaults.notificationsPushEnabled),
    scheduledReportEmail: {
      ...defaults.scheduledReportEmail,
      ...parseObject(data.scheduledReportEmail, defaults.scheduledReportEmail),
      enabled: parseBoolean(data?.scheduledReportEmail?.enabled, defaults.scheduledReportEmail.enabled),
      subject: parseText(data?.scheduledReportEmail?.subject, defaults.scheduledReportEmail.subject),
      body: parseText(data?.scheduledReportEmail?.body, defaults.scheduledReportEmail.body),
      recipients: parseStringArray(data?.scheduledReportEmail?.recipients, defaults.scheduledReportEmail.recipients),
      reportDateMode: String(data?.scheduledReportEmail?.reportDateMode || defaults.scheduledReportEmail.reportDateMode).trim() === 'previous_day'
        ? 'previous_day'
        : defaults.scheduledReportEmail.reportDateMode,
      sendTime: /^\d{2}:\d{2}$/.test(String(data?.scheduledReportEmail?.sendTime || '').trim())
        ? String(data?.scheduledReportEmail?.sendTime || '').trim()
        : defaults.scheduledReportEmail.sendTime
    },
    agentBankOptions: normalizeAgentBankOptions(data.agentBankOptions, defaults.agentBankOptions),
    pfaOptions: parseStringArray([
      ...(defaults.pfaOptions || []),
      ...(Array.isArray(data.pfaOptions) ? data.pfaOptions : [])
    ], defaults.pfaOptions),
    pfaAddresses: parseObject(data.pfaAddresses, defaults.pfaAddresses),
    documentRequirements: normalizeDocumentRequirements(data.documentRequirements, defaults.documentRequirements),
    documentRequirementRoles: normalizeDocumentRequirementRoles(data.documentRequirementRoles, defaults.documentRequirementRoles),
    rolePermissions: {
      ...defaults.rolePermissions,
      ...parseObject(data.rolePermissions, defaults.rolePermissions)
    },
    routingPolicies: {
      ...defaults.routingPolicies,
      ...parseObject(data.routingPolicies, defaults.routingPolicies)
    },
    workflowLabels: {
      ...defaults.workflowLabels,
      ...parseObject(data.workflowLabels, defaults.workflowLabels)
    },
    rejectionRules: {
      ...defaults.rejectionRules,
      ...parseObject(data.rejectionRules, defaults.rejectionRules),
      reviewerCannedReasons: parseStringArray(data?.rejectionRules?.reviewerCannedReasons, defaults.rejectionRules.reviewerCannedReasons),
      rsaCannedReasons: parseStringArray(data?.rejectionRules?.rsaCannedReasons, defaults.rejectionRules.rsaCannedReasons),
      minLength: Math.max(0, parseNumber(data?.rejectionRules?.minLength, defaults.rejectionRules.minLength))
    },
    agentRegistrationRules: {
      ...defaults.agentRegistrationRules,
      ...parseObject(data.agentRegistrationRules, defaults.agentRegistrationRules)
    },
    bulkImportRules: {
      ...defaults.bulkImportRules,
      ...parseObject(data.bulkImportRules, defaults.bulkImportRules),
      requiredColumns: parseStringArray(data?.bulkImportRules?.requiredColumns, defaults.bulkImportRules.requiredColumns)
    },
    uploadControls: {
      ...defaults.uploadControls,
      ...parseObject(data.uploadControls, defaults.uploadControls),
      simulateStorageCapFullFailure: parseBoolean(
        data?.uploadControls?.simulateStorageCapFullFailure,
        defaults.uploadControls.simulateStorageCapFullFailure
      )
    },
    dashboardAnnouncement: {
      ...defaults.dashboardAnnouncement,
      ...parseObject(data.dashboardAnnouncement, defaults.dashboardAnnouncement),
      message: parseText(data?.dashboardAnnouncement?.message, defaults.dashboardAnnouncement.message),
      tone: parseText(data?.dashboardAnnouncement?.tone, defaults.dashboardAnnouncement.tone),
      speed: Math.min(60, Math.max(5, parseNumber(data?.dashboardAnnouncement?.speed, defaults.dashboardAnnouncement.speed))),
      textColor: parseHexColor(data?.dashboardAnnouncement?.textColor, defaults.dashboardAnnouncement.textColor),
      fontSize: Math.min(28, Math.max(12, parseNumber(data?.dashboardAnnouncement?.fontSize, defaults.dashboardAnnouncement.fontSize))),
      fontStyle: ['normal', 'bold', 'italic', 'bold_italic'].includes(String(data?.dashboardAnnouncement?.fontStyle || '').trim().toLowerCase())
        ? String(data.dashboardAnnouncement.fontStyle).trim().toLowerCase()
        : defaults.dashboardAnnouncement.fontStyle,
      fontFamily: ['system', 'arial', 'trebuchet', 'georgia', 'courier', 'verdana', 'tahoma'].includes(String(data?.dashboardAnnouncement?.fontFamily || '').trim().toLowerCase())
        ? String(data.dashboardAnnouncement.fontFamily).trim().toLowerCase()
        : defaults.dashboardAnnouncement.fontFamily,
      targetDashboards: parseAnnouncementTargets(data?.dashboardAnnouncement?.targetDashboards, defaults.dashboardAnnouncement.targetDashboards)
    },
    globalReadOnlyMode: parseBoolean(data.globalReadOnlyMode, defaults.globalReadOnlyMode),
    globalReadOnlyMessage: parseText(data.globalReadOnlyMessage, defaults.globalReadOnlyMessage),
    securityControls: {
      ...defaults.securityControls,
      ...parseObject(data.securityControls, defaults.securityControls),
      sessionTimeoutMinutes: Math.max(1, parseNumber(data?.securityControls?.sessionTimeoutMinutes, defaults.securityControls.sessionTimeoutMinutes)),
      forceLogoutCountdown: parseDurationSetting(
        data?.securityControls?.forceLogoutCountdown ?? data?.securityControls?.forceLogoutCountdownMinutes,
        defaults.securityControls.forceLogoutCountdown
      ),
      forceLogoutToken: parseText(data?.securityControls?.forceLogoutToken, defaults.securityControls.forceLogoutToken)
    },
    notificationTemplates: parseObject(data.notificationTemplates, defaults.notificationTemplates),
    auditControls: {
      ...defaults.auditControls,
      ...parseObject(data.auditControls, defaults.auditControls),
      retentionDays: Math.max(1, parseNumber(data?.auditControls?.retentionDays, defaults.auditControls.retentionDays))
    },
    propertyRules: normalizePropertyRules(data.propertyRules, defaults.propertyRules),
    houseNumberRules: normalizeHouseNumberRules(data.houseNumberRules, defaults.houseNumberRules)
  };
}

export async function getSystemSettings(db, { force = false } = {}) {
  const now = Date.now();
  if (!force && settingsCache && (now - settingsFetchedAt) < CACHE_TTL_MS) {
    return settingsCache;
  }

  try {
    const snap = await getDoc(doc(db, 'settings', 'system'));
    const data = snap.exists() ? (snap.data() || {}) : {};
    settingsCache = normalizeSystemSettings(data);
  } catch (_) {
    settingsCache = getDefaultSystemSettings();
  }

  settingsFetchedAt = now;
  return settingsCache;
}

export function clearSystemSettingsCache() {
  settingsCache = null;
  settingsFetchedAt = 0;
}
