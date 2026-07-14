﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿// js/document-uploader.js - COMPLETE FIXED VERSION WITH WORKING FILE SIZE MODALS AND CALCULATIONS
import { auth, db } from './firebase-config.js';
import { BackblazeStorage } from './backblaze-storage.js?v=20260714a';
import { queueViewerAssignmentEmail } from './email-alerts.js';
import { notifyStatusChangePush } from './status-push.js';
import { notifyAdminPushEvent } from './push-alerts.js';
import { ACCOUNT_LOOKUP_API_BASE_URL, EMAIL_API_BASE_URL } from './email-api-config.js?v=20260703a';
import { formatAppDateTime, getTrustedDateKey, getTrustedNowIso } from './shared/app-time.js';
import {
  getCurrentUserProfile as getCurrentUserProfileShared,
  ensureUserFullNames as ensureUserFullNamesShared,
  isActiveUserWithRole as isActiveUserWithRoleShared,
  getUserFullName as getUserFullNameShared,
  getUserProfileByEmail as getUserProfileByEmailShared,
  normalizeEmail as normalizeEmailShared
} from './shared/user-directory.js?v=20260518a';
import {
  assignRoundRobin as assignRoundRobinShared,
  getUploaderRoutingRule as getUploaderRoutingRuleShared,
  getViewerEmails as getViewerEmailsShared,
  routingRuleDocId as routingRuleDocIdShared
} from './shared/uploader-routing.js?v=20260427e';
import {
  buildSubmissionCommissionFields,
  formatCommissionRateLabel,
  getCommissionSettings,
  getSubmissionCommissionAmount,
  resolveSubmissionCommissionRate
} from './shared/commission-config.js?v=20260507a';
import {
  getTimestampMillis as getStageTimestampMillis,
  getSubmissionCurrentStageEntryAt,
  getSubmissionReviewEntryAt,
  getSubmissionApprovalEntryAt,
  getSubmissionRejectionEntryAt,
  getSubmissionPaymentEntryAt,
  getSubmissionPaidEntryAt,
  getSubmissionClearedEntryAt
} from './shared/submission-stage.js?v=20260609a';
import { getDefaultSystemSettings, getSystemSettings } from './shared/system-settings.js?v=20260617a';
import {
  collection, query, where, orderBy, onSnapshot, addDoc, updateDoc, doc,
  serverTimestamp, arrayUnion, getDocs, getDoc, setDoc, runTransaction, deleteDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

// ==================== FIX JSPDF DETECTION - ADD THIS RIGHT AFTER IMPORTS ====================
(function fixJsPDF() {
    // Check all possible locations where jsPDF might be
    if (window.jspdf && window.jspdf.jsPDF) {
        window.jsPDF = window.jspdf.jsPDF;
    }
    else if (window.jsPDF && typeof window.jsPDF === 'function') {
        window.jspdf = window.jspdf || {};
        window.jspdf.jsPDF = window.jsPDF;
    }
    else if (window.jspdf && typeof window.jspdf === 'function') {
        window.jsPDF = window.jspdf;
        window.jspdf = { jsPDF: window.jspdf };
    }
    else {
        // Create a mock that shows a helpful error message
        window.jspdf = {
            jsPDF: function() {
                throw new Error('jsPDF library not loaded. Please add: <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>');
            }
        };
    }
})();

(function configurePdfJs() {
    if (window.pdfjsLib && !window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    }
})();

// ==================== DOCUMENT TYPES ====================
const DEFAULT_DOCUMENT_TYPES = [
  { id: 'birth_certificate', name: 'Birth Certificate / Age Declaration', icon: 'fa-id-card', required: true },
  { id: 'nin', name: 'National Identification Number (NIN)', icon: 'fa-id-card', required: true },
  { id: 'bvn', name: 'BVN', icon: 'fa-id-badge', required: true },
  { id: 'pay_slips', name: '3 Months Pay Slip', icon: 'fa-file-invoice', required: false },
  { id: 'offer_letter', name: 'Offer of Employment Letter', icon: 'fa-file-signature', required: true },
  { id: 'intro_letter', name: 'Introduction Letter', icon: 'fa-file-signature', required: true },
  { id: 'request_letter', name: 'Request Letter', icon: 'fa-file-signature', required: true },
  { id: 'rsa_statement', name: 'RSA Statement', icon: 'fa-file-invoice', required: true },
  { id: 'pfa_form', name: 'PFA Application Form', icon: 'fa-file-pdf', required: true },
  { id: 'consent_letter', name: 'Consent Letter', icon: 'fa-file-signature', required: true },
  { id: 'indemnity_form', name: 'Indemnity Form', icon: 'fa-file-signature', required: true },
  { id: 'mortgage_loan_application', name: 'Mortgage Loan Application Form', icon: 'fa-file-pdf', required: true },
  { id: 'allocation_last_page', name: 'Allocation Last Page', icon: 'fa-file-pdf', required: true },
  { id: 'offer_letter_last_page', name: 'Offer Letter Last Page', icon: 'fa-file-pdf', required: true },
  { id: 'pmi_soa', name: 'PMI SOA', icon: 'fa-file-pdf', required: true },
  { id: 'benefit_application_form', name: 'Benefit Application Form', icon: 'fa-file-alt', required: false },
  { id: 'data_recapture', name: 'Data Recapture', icon: 'fa-file-alt', required: false },
  { id: 'credit_life', name: 'Credit Life', icon: 'fa-file-medical', required: false }
];

// Make document types globally available
let DOCUMENT_TYPES = [...DEFAULT_DOCUMENT_TYPES];
let REQUIRED_DOC_TYPES = DOCUMENT_TYPES.filter(d => d.required !== false);
let OPTIONAL_DOC_TYPES = DOCUMENT_TYPES.filter(d => d.required === false);

window.REQUIRED_DOC_TYPES = REQUIRED_DOC_TYPES;
window.OPTIONAL_DOC_TYPES = OPTIONAL_DOC_TYPES;

function formatCurrency(value) {
  const num = Number(value || 0);
  try {
    return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
  } catch (e) {
    return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }
}

function roundDownToNearestThousand(value) {
  const num = Number(value || 0);
  return Math.max(0, Math.floor(num / 1000) * 1000);
}

function roundUpToNearestThousand(value) {
  const num = Number(value || 0);
  return Math.max(0, Math.ceil(num / 1000) * 1000);
}

function calculateRoundedRsa25(rsaBalance) {
  return roundDownToNearestThousand((Number(rsaBalance) || 0) * 0.25);
}

function parseMoney(value) {
  const raw = String(value ?? '').replace(/[^0-9.\-]/g, '');
  const num = Number(raw);
  return Number.isFinite(num) ? num : 0;
}

function normalizeLoanAmountValue(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  return roundUpToNearestThousand(parseMoney(raw));
}

function getFinancials(submission) {
  const details = submission?.customerDetails || {};
  const rsaBalance = parseMoney(details.rsaBalance || submission?.rsaBalance || 0);
  const computed25 = calculateRoundedRsa25(rsaBalance);
  const stored25 = parseMoney(details.rsa25Percent || submission?.rsa25Percent || 0);
  const twentyFive = stored25 ? roundDownToNearestThousand(stored25) : computed25;
  const commissionRate = resolveSubmissionCommissionRate(submission);
  const commission2 = getSubmissionCommissionAmount(submission, twentyFive);
  const pfa = String(details.pfa || submission?.pfa || '').trim() || '-';
  return { pfa, twentyFive, commission2, commissionRate, commissionRateLabel: formatCommissionRateLabel(commissionRate) };
}

// ==================== GLOBAL VARIABLES ====================
let currentUser = null;
let currentUserProfile = null;
let allSubmissions = [];
let submissionListenerUnsubs = [];
const submissionSnapshotSources = new Map();
let currentCustomerUploads = {};
let currentEditId = null;
let currentDraftId = null;
let currentHouseNumberReserved = false;
let currentDocType = null;
let currentFile = null;
let approvedAgents = [];
let currentSubmissionAgentFallback = null;
let singlePreviewObjectUrl = null;
let currentDocumentRequirementRoles = { ...(getDefaultSystemSettings().documentRequirementRoles || {}) };
let trustedNowCache = { value: null, fetchedAt: 0 };
// Mobile camera photos can be large; compress only when needed.
let MAX_IMAGE_UPLOAD_BYTES = 1024 * 1024; // 1MB
let MAX_PDF_SIZE_BYTES = 1.5 * 1024 * 1024; // 1.5MB
const userFullNames = new Map();
let customerDetailsSaved = false;
const RR_COUNTER_DOC = doc(db, 'counters', 'roundRobin');
const DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE = '90089';
const DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME = 'Coop Savings and Loans / Cooperative Mortgage Bank';
const FALLBACK_CUSTOMER_ACCOUNT_BANKS = [
  { name: 'Access Bank', code: '044', slug: 'access-bank' },
  { name: 'Access Bank (Diamond)', code: '063', slug: 'access-bank-diamond' },
  { name: 'ALAT by WEMA', code: '035A', slug: 'alat-by-wema' },
  { name: 'Citibank Nigeria', code: '023', slug: 'citibank-nigeria' },
  { name: DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME, code: DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE, slug: 'cooperative-mortgage-bank-ng' },
  { name: 'Ecobank Nigeria', code: '050', slug: 'ecobank-nigeria' },
  { name: 'Fidelity Bank', code: '070', slug: 'fidelity-bank' },
  { name: 'First Bank of Nigeria', code: '011', slug: 'first-bank-of-nigeria' },
  { name: 'First City Monument Bank', code: '214', slug: 'first-city-monument-bank' },
  { name: 'Globus Bank', code: '00103', slug: 'globus-bank' },
  { name: 'Guaranty Trust Bank', code: '058', slug: 'guaranty-trust-bank' },
  { name: 'Keystone Bank', code: '082', slug: 'keystone-bank' },
  { name: 'Kuda Bank', code: '50211', slug: 'kuda-bank' },
  { name: 'Lotus Bank', code: '303', slug: 'lotus-bank' },
  { name: 'Moniepoint MFB', code: '50515', slug: 'moniepoint-mfb-ng' },
  { name: 'OPay Digital Services Limited (OPay)', code: '999992', slug: 'paycom' },
  { name: 'Optimus Bank Limited', code: '107', slug: 'optimus-bank-ltd' },
  { name: 'PalmPay', code: '999991', slug: 'palmpay' },
  { name: 'Parallex Bank', code: '104', slug: 'parallex-bank' },
  { name: 'Polaris Bank', code: '076', slug: 'polaris-bank' },
  { name: 'PremiumTrust Bank', code: '105', slug: 'premiumtrust-bank-ng' },
  { name: 'Providus Bank', code: '101', slug: 'providus-bank' },
  { name: 'Rubies MFB', code: '125', slug: 'rubies-mfb' },
  { name: 'Stanbic IBTC Bank', code: '221', slug: 'stanbic-ibtc-bank' },
  { name: 'Standard Chartered Bank', code: '068', slug: 'standard-chartered-bank' },
  { name: 'Sterling Bank', code: '232', slug: 'sterling-bank' },
  { name: 'Suntrust Bank', code: '100', slug: 'suntrust-bank' },
  { name: 'Titan Bank', code: '102', slug: 'titan-bank' },
  { name: 'Union Bank of Nigeria', code: '032', slug: 'union-bank-of-nigeria' },
  { name: 'United Bank For Africa', code: '033', slug: 'united-bank-for-africa' },
  { name: 'Unity Bank', code: '215', slug: 'unity-bank' },
  { name: 'VFD Microfinance Bank Limited', code: '566', slug: 'vfd' },
  { name: 'Wema Bank', code: '035', slug: 'wema-bank' },
  { name: 'Zenith Bank', code: '057', slug: 'zenith-bank' }
];
function formatUploadLimitLabel(bytes) {
  const mb = Number(bytes || 0) / (1024 * 1024);
  if (!Number.isFinite(mb) || mb <= 0) return '0 MB';
  return `${mb.toFixed(mb >= 10 ? 0 : 2).replace(/\.00$/, '')}MB`;
}

function getImageUploadLimitLabel() {
  return formatUploadLimitLabel(MAX_IMAGE_UPLOAD_BYTES);
}

function getPdfUploadLimitLabel() {
  return formatUploadLimitLabel(MAX_PDF_SIZE_BYTES);
}

function assertWritable(actionLabel) {
  return typeof window.assertAppWritable === 'function'
    ? window.assertAppWritable(actionLabel)
    : true;
}

async function applyUploadSystemSettings() {
  try {
    const settings = await getSystemSettings(db);
    MAX_IMAGE_UPLOAD_BYTES = Math.max(0.1, Number(settings.maxImageUploadMb || 1)) * 1024 * 1024;
    MAX_PDF_SIZE_BYTES = Math.max(0.1, Number(settings.maxPdfUploadMb || 1.5)) * 1024 * 1024;
  } catch (_) {
    MAX_IMAGE_UPLOAD_BYTES = 1024 * 1024;
    MAX_PDF_SIZE_BYTES = 1.5 * 1024 * 1024;
  }
}

async function populateAgentBankOptions({ force = false } = {}) {
  if (!agentAccountBankSelect) return;

  const currentValue = String(agentAccountBankSelect.value || '').trim();
  let bankOptions = !force && accountLookupBanks.length ? [...accountLookupBanks] : [];
  try {
    if (!bankOptions.length) {
      bankOptions = await fetchAccountLookupBanks();
    }
  } catch (_) {
    bankOptions = [...FALLBACK_CUSTOMER_ACCOUNT_BANKS];
  }
  if (!bankOptions.length) bankOptions = [...FALLBACK_CUSTOMER_ACCOUNT_BANKS];
  agentAccountBankOptions = [...bankOptions];

  if (agentAccountBankSelect.tagName === 'SELECT') {
    agentAccountBankSelect.innerHTML = '<option value="">Select Bank</option>' + bankOptions.map((bank) => (
      `<option value="${escapeHtml(bank.code || bank.name)}" data-name="${escapeHtml(getCustomerAccountBankLabel(bank))}">${escapeHtml(getCustomerAccountBankLabel(bank))}</option>`
    )).join('');
  } else {
    renderAgentBankDatalist(bankOptions);
  }

  if (currentValue) {
    const currentBank = findAgentAccountBankByValue(currentValue, bankOptions);
    if (currentBank) {
      agentAccountBankSelect.value = agentAccountBankSelect.tagName === 'SELECT'
        ? (currentBank.code || currentBank.name)
        : getCustomerAccountBankLabel(currentBank);
    }
  }
}

function getBackendApiBaseUrl() {
  const runtime = String(window.__ACCOUNT_LOOKUP_API_BASE_URL__ || '').trim();
  const configured = runtime
    || String(ACCOUNT_LOOKUP_API_BASE_URL || '').trim()
    || String(window.__EMAIL_API_BASE_URL__ || '').trim()
    || String(EMAIL_API_BASE_URL || '').trim();
  if (!configured || configured.includes('YOUR-RENDER-URL')) return '';
  return configured.replace(/\/+$/, '');
}

async function fetchAccountLookupBanks() {
  const data = await backendFetchJson('/api/paystack/banks');
  return Array.isArray(data.banks) ? data.banks : [];
}

async function fetchResolvedAccountName(accountNumber, bankCode) {
  const data = await backendFetchJson('/api/paystack/resolve-account', {
    method: 'POST',
    body: JSON.stringify({ accountNumber, bankCode })
  });
  return String(data.accountName || '').trim();
}

function setAccountLookupStatus(message = '', type = 'info') {
  if (!accountLookupStatus) return;
  const text = String(message || '').trim();
  accountLookupStatus.textContent = text;
  accountLookupStatus.style.display = text ? 'block' : 'none';
  const colors = {
    error: '#dc2626',
    success: '#15803d',
    info: '#64748b'
  };
  accountLookupStatus.style.color = colors[type] || colors.info;
  if (accountNoInput) {
    accountNoInput.style.borderColor = type === 'error' ? '#dc2626' : '';
    accountNoInput.style.boxShadow = type === 'error' ? '0 0 0 3px rgba(220, 38, 38, 0.12)' : '';
  }
  if (accountBankSelect) {
    accountBankSelect.style.borderColor = type === 'error' ? '#dc2626' : '';
    accountBankSelect.style.boxShadow = type === 'error' ? '0 0 0 3px rgba(220, 38, 38, 0.12)' : '';
  }
}

function clearVerifiedAccountLookup() {
  verifiedAccountLookup = { accountNumber: '', bankCode: '', accountName: '' };
  if (accountNameInput) accountNameInput.value = '';
  if (customerNameInput) customerNameInput.value = '';
}

function getSelectedAccountBank() {
  if (!accountBankSelect || accountBankSelect.tagName !== 'SELECT') {
    if (accountBankSelect) accountBankSelect.value = DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME;
    return { code: DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE, name: DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME };
  }
  const code = String(accountBankSelect?.value || '').trim();
  const option = accountBankSelect?.selectedOptions?.[0] || null;
  const name = String(option?.dataset?.name || option?.textContent || '').trim();
  return {
    code: code || DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE,
    name: name || DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME
  };
}

function getCustomerAccountBankLabel(bank = {}) {
  const code = String(bank.code || '').trim();
  const slug = String(bank.slug || '').trim();
  if (code === '90089' || slug === 'cooperative-mortgage-bank-ng') {
    return 'Coop Savings and Loans / Cooperative Mortgage Bank';
  }
  return String(bank.name || '').trim();
}

function renderCustomerAccountBankOptions(banks = [], currentCode = '') {
  if (!accountBankSelect) return;
  if (accountBankSelect.tagName !== 'SELECT') {
    accountBankSelect.value = DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME;
    return;
  }
  const options = Array.isArray(banks) ? banks : [];
  accountBankSelect.innerHTML = '<option value="">Select Bank</option>' + options.map((bank) => (
    `<option value="${escapeHtml(bank.code)}" data-name="${escapeHtml(getCustomerAccountBankLabel(bank))}">${escapeHtml(getCustomerAccountBankLabel(bank))}</option>`
  )).join('');
  const selectedCode = currentCode || DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE;
  if (selectedCode && options.some((bank) => bank.code === selectedCode)) {
    accountBankSelect.value = selectedCode;
  }
}

async function backendFetchJson(path, options = {}) {
  const baseUrl = getBackendApiBaseUrl();
  if (!baseUrl) throw new Error('Account lookup service is not configured');
  const idToken = await auth.currentUser?.getIdToken?.();
  if (!idToken) throw new Error('Please sign in again to verify account details');
  const response = await fetch(`${baseUrl}${path}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${idToken}`,
      ...(options.headers || {})
    }
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.ok === false) {
    if (response.status === 404 && String(path || '').startsWith('/api/paystack/')) {
      throw new Error('Account lookup endpoint is not deployed on the backend yet.');
    }
    throw new Error(String(data.error || data.message || `Request failed with ${response.status}`));
  }
  return data;
}

async function populateCustomerAccountBankOptions() {
  if (!accountBankSelect) return;
  if (accountBankSelect.tagName !== 'SELECT') {
    accountLookupBanks = [...FALLBACK_CUSTOMER_ACCOUNT_BANKS];
    accountBankSelect.value = DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME;
    setAccountLookupStatus('', 'info');
    return;
  }
  const currentCode = String(accountBankSelect.value || '').trim();
  try {
    accountLookupBanks = await fetchAccountLookupBanks();
    if (!accountLookupBanks.length) {
      accountLookupBanks = [...FALLBACK_CUSTOMER_ACCOUNT_BANKS];
      setAccountLookupStatus('Using saved bank list. Account name verification still needs Paystack.', 'info');
    } else {
      setAccountLookupStatus('', 'info');
    }
    renderCustomerAccountBankOptions(accountLookupBanks, currentCode);
  } catch (error) {
    accountLookupBanks = [...FALLBACK_CUSTOMER_ACCOUNT_BANKS];
    renderCustomerAccountBankOptions(accountLookupBanks, currentCode);
    setAccountLookupStatus(
      `${error.message || 'Bank list unavailable'} Using saved bank list for now.`,
      'error'
    );
  }
}

async function resolveCustomerAccountName({ silent = false } = {}) {
  const sequence = ++accountLookupSequence;
  const accountNumber = String(accountNoInput?.value || '').replace(/\D/g, '');
  const bank = getSelectedAccountBank();
  clearVerifiedAccountLookup();

  if (!accountNumber && !bank.code) {
    setAccountLookupStatus('', 'info');
    return true;
  }
  if (accountNumber && accountNumber.length !== 10) {
    setAccountLookupStatus('Enter a valid 10-digit account number.', 'error');
    return false;
  }
  if (accountNumber.length === 10 && !bank.code) {
    setAccountLookupStatus('Select account bank to verify account name.', 'error');
    return false;
  }
  if (accountNumber.length !== 10 || !bank.code) return true;

  if (!silent) setAccountLookupStatus('Verifying account name...', 'info');
  try {
    const accountName = await fetchResolvedAccountName(accountNumber, bank.code);
    if (sequence !== accountLookupSequence) return true;
    verifiedAccountLookup = { accountNumber, bankCode: bank.code, accountName };
    if (accountNameInput) accountNameInput.value = accountName;
    if (customerNameInput) {
      customerNameInput.value = accountName;
      updateSubmitButton();
    }
    setAccountLookupStatus(accountName ? `Verified: ${accountName}` : 'Account verified.', 'success');
    return true;
  } catch (error) {
    if (sequence !== accountLookupSequence) return true;
    setAccountLookupStatus(error.message || 'Could not verify account name.', 'error');
    return false;
  }
}

function scheduleCustomerAccountLookup() {
  if (accountLookupDebounce) window.clearTimeout(accountLookupDebounce);
  accountLookupDebounce = window.setTimeout(() => {
    void resolveCustomerAccountName();
  }, 500);
}

function setAgentAccountLookupStatus(message = '', type = 'info') {
  if (!agentAccountLookupStatus) return;
  const text = String(message || '').trim();
  agentAccountLookupStatus.textContent = text;
  agentAccountLookupStatus.style.display = text ? 'block' : 'none';
  const colors = {
    error: '#dc2626',
    success: '#15803d',
    info: '#64748b'
  };
  agentAccountLookupStatus.style.color = colors[type] || colors.info;
  if (agentAccountNumberInput) {
    agentAccountNumberInput.style.borderColor = type === 'error' ? '#dc2626' : '';
    agentAccountNumberInput.style.boxShadow = type === 'error' ? '0 0 0 3px rgba(220, 38, 38, 0.12)' : '';
  }
  if (agentAccountBankSelect) {
    agentAccountBankSelect.style.borderColor = type === 'error' ? '#dc2626' : '';
    agentAccountBankSelect.style.boxShadow = type === 'error' ? '0 0 0 3px rgba(220, 38, 38, 0.12)' : '';
  }
}

function clearVerifiedAgentAccountLookup() {
  verifiedAgentAccountLookup = { accountNumber: '', bankCode: '', accountName: '' };
  if (agentAccountNameInput) agentAccountNameInput.value = '';
}

function findAgentAccountBankByValue(value = '', banks = null) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  const normalized = raw.toLowerCase();
  const options = Array.isArray(banks) && banks.length
    ? banks
    : (agentAccountBankOptions.length ? agentAccountBankOptions : [...accountLookupBanks, ...FALLBACK_CUSTOMER_ACCOUNT_BANKS]);
  return options.find((bank) => {
    const label = getCustomerAccountBankLabel(bank);
    return [bank.code, bank.name, label].some((candidate) => String(candidate || '').trim().toLowerCase() === normalized);
  }) || null;
}

function getSelectedAgentAccountBank() {
  if (!agentAccountBankSelect) return { code: '', name: '' };
  if (agentAccountBankSelect.tagName !== 'SELECT') {
    const value = String(agentAccountBankSelect.value || '').trim();
    const bank = findAgentAccountBankByValue(value);
    return bank
      ? { code: String(bank.code || '').trim(), name: getCustomerAccountBankLabel(bank) }
      : { code: '', name: value };
  }
  const code = String(agentAccountBankSelect?.value || '').trim();
  const option = agentAccountBankSelect?.selectedOptions?.[0] || null;
  const name = String(option?.dataset?.name || option?.textContent || '').trim();
  return { code, name };
}

async function resolveAgentAccountName({ silent = false } = {}) {
  const sequence = ++agentAccountLookupSequence;
  const accountNumber = String(agentAccountNumberInput?.value || '').replace(/\D/g, '');
  const bank = getSelectedAgentAccountBank();
  clearVerifiedAgentAccountLookup();

  if (!accountNumber && !bank.code) {
    setAgentAccountLookupStatus('', 'info');
    return true;
  }
  if (accountNumber && accountNumber.length !== 10) {
    setAgentAccountLookupStatus('Enter a valid 10-digit account number.', 'error');
    return false;
  }
  if (accountNumber.length === 10 && !bank.code) {
    setAgentAccountLookupStatus('Select account bank to verify account name.', 'error');
    return false;
  }
  if (accountNumber.length !== 10 || !bank.code) return true;

  if (!silent) setAgentAccountLookupStatus('Verifying account name...', 'info');
  try {
    const accountName = await fetchResolvedAccountName(accountNumber, bank.code);
    if (sequence !== agentAccountLookupSequence) return true;
    verifiedAgentAccountLookup = { accountNumber, bankCode: bank.code, accountName };
    if (agentAccountNameInput) agentAccountNameInput.value = accountName;
    setAgentAccountLookupStatus(accountName ? `Verified: ${accountName}` : 'Account verified.', 'success');
    return true;
  } catch (error) {
    if (sequence !== agentAccountLookupSequence) return true;
    setAgentAccountLookupStatus(error.message || 'Could not verify account name.', 'error');
    return false;
  }
}

function scheduleAgentAccountLookup() {
  if (agentAccountLookupDebounce) window.clearTimeout(agentAccountLookupDebounce);
  agentAccountLookupDebounce = window.setTimeout(() => {
    void resolveAgentAccountName();
  }, 500);
}

function renderAgentBankDatalist(banks = []) {
  const options = Array.isArray(banks) && banks.length ? banks : agentAccountBankOptions;
  const list = document.getElementById('agentAccountBankList');
  if (!list) return;
  list.innerHTML = options.map((bank) => {
    const label = getCustomerAccountBankLabel(bank);
    const accountName = String(bank.accountName || '').trim();
    return `<option value="${escapeHtml(label)}" label="${escapeHtml(accountName ? `${label} - ${accountName}` : label)}"></option>`;
  }).join('');
}

function applyDocumentRequirements(documentRequirements = []) {
  const activeDocs = Array.isArray(documentRequirements) && documentRequirements.length
    ? documentRequirements.filter((doc) => doc?.active !== false)
    : DEFAULT_DOCUMENT_TYPES;

  DOCUMENT_TYPES = activeDocs.map((doc) => ({
    id: String(doc.id || '').trim(),
    name: String(doc.name || '').trim(),
    icon: String(doc.icon || 'fa-file-alt').trim(),
    required: doc.required !== false
  })).filter((doc) => doc.id && doc.name);

  REQUIRED_DOC_TYPES = DOCUMENT_TYPES.filter((doc) => doc.required !== false);
  OPTIONAL_DOC_TYPES = DOCUMENT_TYPES.filter((doc) => doc.required === false);
  window.REQUIRED_DOC_TYPES = REQUIRED_DOC_TYPES;
  window.OPTIONAL_DOC_TYPES = OPTIONAL_DOC_TYPES;
  if (totalCountSpan) totalCountSpan.textContent = String(DOCUMENT_TYPES.length);
}

function populatePfaOptions(pfaOptions = []) {
  if (!pfaOptionsList) return;
  const options = Array.isArray(pfaOptions) && pfaOptions.length ? pfaOptions : [];
  pfaOptionsList.innerHTML = options.map((pfa) => `<option value="${escapeHtml(pfa)}"></option>`).join('');
  if (pfaInput && !pfaInput.getAttribute('list')) {
    pfaInput.setAttribute('list', 'pfaOptionsList');
  }
}

async function applyWorkflowSystemSettings({ force = false } = {}) {
  try {
    const settings = await getSystemSettings(db, { force });
    applyDocumentRequirements(settings.documentRequirements);
    currentDocumentRequirementRoles = {
      ...(getDefaultSystemSettings().documentRequirementRoles || {}),
      ...(settings.documentRequirementRoles || {})
    };
    populatePfaOptions(settings.pfaOptions);
    PROPERTY_RULES = Array.isArray(settings.propertyRules) && settings.propertyRules.length ? [...settings.propertyRules] : [...DEFAULT_PROPERTY_RULES];
    HOUSE_NUMBER_RULES = settings.houseNumberRules && typeof settings.houseNumberRules === 'object'
      ? { ...settings.houseNumberRules }
      : { ...DEFAULT_HOUSE_NUMBER_RULES };
  } catch (_) {
    applyDocumentRequirements(DEFAULT_DOCUMENT_TYPES);
    currentDocumentRequirementRoles = { ...(getDefaultSystemSettings().documentRequirementRoles || {}) };
    populatePfaOptions([]);
    PROPERTY_RULES = [...DEFAULT_PROPERTY_RULES];
    HOUSE_NUMBER_RULES = { ...DEFAULT_HOUSE_NUMBER_RULES };
  }
  syncUploadRequirementUi();
  updateSubmitButton();
}

// ==================== INJECT MODAL ANIMATIONS (GLOBAL) ====================
function injectModalAnimations() {
  if (document.getElementById('modalAnimations')) return;
  const style = document.createElement('style');
  style.id = 'modalAnimations';
  style.textContent = `
    @keyframes modalSlideIn {
      from { transform: translateY(-30px); opacity: 0; }
      to { transform: translateY(0); opacity: 1; }
    }
  `;
  document.head.appendChild(style);
}
// Call on script load
injectModalAnimations();

// ==================== FILE SIZE WARNING MODAL - SINGLE UPLOAD ====================
function showFileSizeWarningModal(file) {
  // Remove any existing modals first
  const existingModal = document.getElementById('fileSizeWarningModal');
  if (existingModal) {
    existingModal.remove();
  }

  const formatSize = (bytes) => (bytes / (1024 * 1024)).toFixed(2) + ' MB';

  // Create modal element
  const modal = document.createElement('div');
  modal.className = 'modal file-size-warning';
  modal.id = 'fileSizeWarningModal';

  // CRITICAL: Use !important to override CSS conflicts
  modal.style.cssText = `
    display: flex !important;
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    width: 100% !important;
    height: 100% !important;
    background: rgba(0, 0, 0, 0.5) !important;
    z-index: 999999 !important;
    align-items: center !important;
    justify-content: center !important;
    opacity: 1 !important;
  `;

  const modalContent = document.createElement('div');
  modalContent.style.cssText = `
    background: white !important;
    border-radius: 12px !important;
    max-width: 500px !important;
    width: 90% !important;
    max-height: 90vh !important;
    overflow-y: auto !important;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25) !important;
    animation: modalSlideIn 0.3s ease !important;
  `;

  modalContent.innerHTML = `
    <div style="padding: 20px; border-bottom: 1px solid #e2e8f0; background: #fee2e2; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
      <h2 style="margin: 0; color: #991b1b; font-size: 20px; display: flex; align-items: center; gap: 10px;">
        <i class="fas fa-exclamation-triangle"></i> File Too Large
      </h2>
      <button id="fileSizeWarningCloseBtn" style="background: transparent; border: none; font-size: 28px; color: #991b1b; cursor: pointer; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; border-radius: 50%; transition: all 0.2s;">&times;</button>
    </div>
    <div style="padding: 25px;">
      <div style="text-align: center; margin-bottom: 20px;">
        <i class="fas fa-file-pdf" style="font-size: 50px; color: #ef4444;"></i>
      </div>
      <p style="font-size: 16px; color: #1e293b; margin-bottom: 15px; text-align: center;">
        <strong>Maximum file size allowed is ${getPdfUploadLimitLabel()}</strong>
      </p>
      <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin-bottom: 20px; text-align: center;">
        <p style="font-size: 16px; font-weight: 600; color: #1e293b; margin-bottom: 10px;">${file.name}</p>
        <p style="font-size: 18px; color: #ef4444; font-weight: 700; background: #fee2e2; display: inline-block; padding: 8px 16px; border-radius: 20px;">
          Size: ${formatSize(file.size)}
        </p>
      </div>
      <div style="background: #f1f5f9; border-radius: 8px; padding: 15px;">
        <p style="font-weight: 600; margin-bottom: 10px; color: #334155;">
          <i class="fas fa-lightbulb" style="color: #f59e0b;"></i> Tips to reduce file size:
        </p>
        <ul style="padding-left: 20px; color: #475569; font-size: 13px; line-height: 1.6;">
          <li>Compress PDF using online tools (ilovepdf.com, smallpdf.com)</li>
          <li>Reduce image resolution before converting to PDF</li>
          <li>Scan documents at 150 DPI instead of 300 DPI</li>
          <li>Save images as JPG instead of PNG</li>
        </ul>
      </div>
    </div>
    <div style="padding: 20px; border-top: 1px solid #e2e8f0; display: flex; justify-content: center; background: #f8fafc; border-radius: 0 0 12px 12px;">
      <button id="understandBtn" style="background: #003366; color: white; border: none; padding: 12px 30px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 8px;">
        <i class="fas fa-check"></i> Understood
      </button>
    </div>
  `;

  modal.appendChild(modalContent);

  // SAFER: Append to body or html fallback
  const target = document.body || document.documentElement;
  if (!target) {
    return;
  }
  target.appendChild(modal);

  // Add event listeners
  const closeBtn = document.getElementById('fileSizeWarningCloseBtn');
  const understandBtn = document.getElementById('understandBtn');

  if (closeBtn) closeBtn.addEventListener('click', () => { modal.remove(); });
  if (understandBtn) understandBtn.addEventListener('click', () => { modal.remove(); });

  modal.addEventListener('click', (e) => {
    if (e.target === modal) { modal.remove(); }
  });
}

// ==================== BATCH SIZE WARNING MODAL ====================
function showBatchSizeWarningModal(oversizedFiles, validFiles, originalFiles) {
  const existingModal = document.getElementById('batchSizeWarningModal');
  if (existingModal) existingModal.remove();

  const formatSize = (bytes) => (bytes / (1024 * 1024)).toFixed(2) + ' MB';

  const modal = document.createElement('div');
  modal.className = 'modal batch-warning';
  modal.id = 'batchSizeWarningModal';

  modal.style.cssText = `
    display: flex !important;
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    width: 100% !important;
    height: 100% !important;
    background: rgba(0, 0, 0, 0.5) !important;
    z-index: 999999 !important;
    align-items: center !important;
    justify-content: center !important;
    opacity: 1 !important;
  `;

  const modalContent = document.createElement('div');
  modalContent.style.cssText = `
    background: white !important;
    border-radius: 12px !important;
    max-width: 700px !important;
    width: 95% !important;
    max-height: 90vh !important;
    overflow-y: auto !important;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25) !important;
    animation: modalSlideIn 0.3s ease !important;
  `;

  const oversizedListHtml = oversizedFiles.map(file => `
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px; background: #fef2f2; border-bottom: 1px solid #fecaca;">
      <div style="display: flex; align-items: center; gap: 10px; flex: 1;">
        <i class="fas fa-exclamation-circle" style="color: #ef4444;"></i>
        <span style="font-weight: 500; color: #1e293b;">${file.name}</span>
      </div>
      <span style="color: #ef4444; font-weight: 600; background: #fee2e2; padding: 4px 8px; border-radius: 4px; white-space: nowrap;">
        ${formatSize(file.size)}
      </span>
    </div>
  `).join('');

  const validListHtml = validFiles.map(file => `
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px; background: #f0fdf4; border-bottom: 1px solid #bbf7d0;">
      <div style="display: flex; align-items: center; gap: 10px; flex: 1;">
        <i class="fas fa-check-circle" style="color: #10b981;"></i>
        <span style="font-weight: 500; color: #1e293b;">${file.name}</span>
      </div>
      <span style="color: #10b981; font-weight: 600; background: #dcfce7; padding: 4px 8px; border-radius: 4px; white-space: nowrap;">
        ${formatSize(file.size)}
      </span>
    </div>
  `).join('');

  modalContent.innerHTML = `
    <div style="padding: 20px; border-bottom: 1px solid #e2e8f0; background: #fffbeb; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
      <h2 style="margin: 0; color: #92400e; font-size: 20px; display: flex; align-items: center; gap: 10px;">
        <i class="fas fa-exclamation-triangle"></i> File Size Validation
      </h2>
      <button id="batchSizeWarningCloseBtn" style="background: transparent; border: none; font-size: 28px; color: #92400e; cursor: pointer; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; border-radius: 50%; transition: all 0.2s;">&times;</button>
    </div>
    <div style="padding: 25px; max-height: 60vh; overflow-y: auto;">
      <p style="font-size: 16px; color: #1e293b; margin-bottom: 20px; text-align: center;">
        <strong>${oversizedFiles.length}</strong> file(s) exceed the ${getPdfUploadLimitLabel()} limit.
        <strong>${validFiles.length}</strong> file(s) are already within limit.
      </p>
      ${oversizedFiles.length > 0 ? `
        <div style="margin-bottom: 25px;">
          <h3 style="font-size: 16px; font-weight: 600; color: #991b1b; margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
            <i class="fas fa-compress-alt"></i> Files Exceeding ${getPdfUploadLimitLabel()}
          </h3>
          <div style="border: 1px solid #fecaca; border-radius: 8px; overflow: hidden;">
            ${oversizedListHtml}
          </div>
        </div>
      ` : ''}
      ${validFiles.length > 0 ? `
        <div style="margin-bottom: 20px;">
          <h3 style="font-size: 16px; font-weight: 600; color: #065f46; margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
            <i class="fas fa-check-circle"></i> Files Within Limit (will be uploaded)
          </h3>
          <div style="border: 1px solid #bbf7d0; border-radius: 8px; overflow: hidden;">
            ${validListHtml}
          </div>
        </div>
      ` : ''}
    </div>
    <div style="padding: 20px; border-top: 1px solid #e2e8f0; display: flex; gap: 10px; justify-content: flex-end; background: #f8fafc; border-radius: 0 0 12px 12px;">
      <button id="cancelBtn" style="background: white; border: 1px solid #cbd5e1; color: #475569; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 500; cursor: pointer; display: flex; align-items: center; gap: 8px;">
        <i class="fas fa-times"></i> Cancel All
      </button>
      ${oversizedFiles.length > 0 ? `
        <button id="compressProceedBtn" style="background: #003366; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 8px;">
          <i class="fas fa-compress-alt"></i> Compress Oversized and Continue
        </button>
      ` : ''}
      ${validFiles.length > 0 ? `
        <button id="proceedBtn" style="background: #10b981; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 8px;">
          <i class="fas fa-upload"></i> Upload ${validFiles.length} Valid File(s)
        </button>
      ` : ''}
    </div>
  `;

  modal.appendChild(modalContent);

  const target = document.body || document.documentElement;
  if (!target) {
    return;
  }
  target.appendChild(modal);

  // Event listeners
  const closeBtn = document.getElementById('batchSizeWarningCloseBtn');
  const cancelBtn = document.getElementById('cancelBtn');
  const proceedBtn = document.getElementById('proceedBtn');
  const compressProceedBtn = document.getElementById('compressProceedBtn');

  if (closeBtn) closeBtn.addEventListener('click', () => { modal.remove(); });
  if (cancelBtn) cancelBtn.addEventListener('click', () => { modal.remove(); });

  if (proceedBtn) {
    proceedBtn.addEventListener('click', () => {
      modal.remove();
      if (validFiles.length > 0) prepareBatchForMapping(validFiles);
    });
  }

  if (compressProceedBtn) {
    compressProceedBtn.addEventListener('click', async () => {
      modal.remove();
      const filesToUpload = [...validFiles];
      const failedFiles = [];
      if (window.showLoader) window.showLoader('Compressing oversized files...');
      try {
        for (const file of oversizedFiles) {
          try {
            const compressed = await compressFileToUploadLimit(file);
            if (compressed.size <= MAX_PDF_SIZE_BYTES) {
              filesToUpload.push(compressed);
            } else {
              failedFiles.push(file.name);
            }
          } catch (error) {
            failedFiles.push(file.name);
          }
        }
      } finally {
        hideLoader();
      }
      if (failedFiles.length > 0) {
        showNotification(`Some files could not be compressed enough: ${failedFiles.join(', ')}`, 'warning');
      }
      if (filesToUpload.length > 0) {
        prepareBatchForMapping(filesToUpload);
      } else {
        showNotification('No files are ready for upload after compression.', 'error');
      }
    });
  }

  modal.addEventListener('click', (e) => {
    if (e.target === modal) { modal.remove(); }
  });
}

// ==================== GET VIEWER EMAILS ====================
async function getViewerEmails() {
  return getViewerEmailsShared(db);
}

function normalizeEmail(email) {
  return normalizeEmailShared(email);
}

function routingRuleDocId(uploaderEmail) {
  return routingRuleDocIdShared(uploaderEmail);
}

async function getUploaderRoutingRule(uploaderEmail) {
  return getUploaderRoutingRuleShared(db, uploaderEmail);
}

async function isActiveUserWithRole(email, allowedRoles = []) {
  return isActiveUserWithRoleShared(db, email, allowedRoles);
}

async function assignRoundRobin(subRef) {
  return assignRoundRobinShared({ db, currentUser, subRef, counterDoc: RR_COUNTER_DOC });
}

async function isActiveRSAUser(email, { allowRoundRobinSkipped = false } = {}) {
  const normalized = normalizeEmail(email);
  if (!normalized) return false;
  try {
    const data = await getUserProfileByEmailShared(db, normalized);
    if (!data) return false;
    const role = String(data.role || '').trim().toLowerCase();
    const status = String(data.status || 'active').toLowerCase();
    const leaveStatus = String(data.leaveStatus || '').toLowerCase();
    return role === 'rsa'
      && status !== 'deactivated'
      && leaveStatus !== 'on_leave'
      && (allowRoundRobinSkipped || data?.skipRsaRoundRobin !== true);
  } catch (_) {
    return false;
  }
}

async function markDirectRsaManualReview(subRef, assignmentMode = 'manual_review') {
  await updateDoc(subRef, {
    assignedTo: '',
    assignedToRSA: '',
    rsaAssignedAt: null,
    status: 'processing_to_pfa',
    rsaReady: true,
    assignmentMode: 'skip_reviewer_routing',
    rsaAssignmentMode: assignmentMode,
    reviewerSkipped: true
  });
}

async function getRSAEmails() {
  const snap = await getDocs(collection(db, 'users'));
  return snap.docs
    .map((d) => d.data() || {})
    .filter((u) => String(u.role || '').trim().toLowerCase() === 'rsa')
    .filter((u) => String(u.status || 'active').toLowerCase() !== 'deactivated'
      && String(u.leaveStatus || '').toLowerCase() !== 'on_leave'
      && u?.skipRsaRoundRobin !== true)
    .map((u) => normalizeEmail(u.email))
    .filter(Boolean)
    .sort();
}

function pickFallbackRSAUser(rsaUsers, seed = '') {
  if (!rsaUsers.length) return '';
  const storageKey = 'rsaRoundRobinFallbackIndex';
  try {
    const lastIndex = Number(window.localStorage?.getItem(storageKey) ?? -1);
    const nextIndex = (Number.isFinite(lastIndex) ? lastIndex + 1 : 0) % rsaUsers.length;
    window.localStorage?.setItem(storageKey, String(nextIndex));
    return rsaUsers[nextIndex] || rsaUsers[0] || '';
  } catch (_) {
    const text = String(seed || Date.now());
    let hash = 0;
    for (let i = 0; i < text.length; i += 1) {
      hash = ((hash << 5) - hash + text.charCodeAt(i)) | 0;
    }
    return rsaUsers[Math.abs(hash) % rsaUsers.length] || rsaUsers[0] || '';
  }
}

async function assignDirectToRSA(subRef, uploaderEmail = '') {
  const routingRule = await getUploaderRoutingRule(uploaderEmail);
  const mappedRsa = routingRule?.rsaEmail || '';
  if (mappedRsa && await isActiveRSAUser(mappedRsa, { allowRoundRobinSkipped: true })) {
    await updateDoc(subRef, {
      assignedTo: '',
      assignedToRSA: mappedRsa,
      rsaAssignedAt: serverTimestamp(),
      status: 'processing_to_pfa',
      rsaReady: true,
      assignmentMode: 'skip_reviewer_routing',
      rsaAssignmentMode: 'uploader_routing',
      reviewerSkipped: true
    });
    return mappedRsa;
  }

  const systemSettings = await getSystemSettings(db);
  const fallbackMode = String(systemSettings.routingPolicies?.fallbackAssignmentMode || 'round_robin').trim().toLowerCase();
  const rsaUsers = await getRSAEmails();
  if (!rsaUsers.length) {
    await markDirectRsaManualReview(subRef, 'no_rsa_available');
    return '';
  }
  if (!systemSettings.rsaRoundRobinEnabled || fallbackMode === 'manual_review') {
    await markDirectRsaManualReview(subRef, fallbackMode === 'manual_review' ? 'manual_review' : 'round_robin_disabled');
    return '';
  }

  const counterRef = doc(db, 'counters', 'roundRobinRSA');
  let assigned = '';
  const trustedDateKey = await getTrustedDateKey();
  try {
    await runTransaction(db, async (tx) => {
      let lastIndex = -1;
      const counterSnap = await tx.get(counterRef);
      if (counterSnap.exists()) {
        const data = counterSnap.data() || {};
        lastIndex = typeof data.lastIndex === 'number' ? data.lastIndex : -1;
      }
      const newIndex = (lastIndex + 1) % rsaUsers.length;
      assigned = rsaUsers[newIndex];
      tx.set(counterRef, { lastIndex: newIndex, lastDate: trustedDateKey }, { merge: true });
      tx.update(subRef, {
        assignedTo: '',
        assignedToRSA: assigned,
        rsaAssignedAt: serverTimestamp(),
        status: 'processing_to_pfa',
        rsaReady: true,
        assignmentMode: 'skip_reviewer_routing',
        rsaAssignmentMode: 'round_robin',
        reviewerSkipped: true
      });
    });
  } catch (error) {
    if (fallbackMode === 'manual_review') {
      await markDirectRsaManualReview(subRef, 'manual_review');
    } else {
      assigned = pickFallbackRSAUser(rsaUsers, subRef?.id || trustedDateKey);
      if (assigned) {
        await updateDoc(subRef, {
          assignedTo: '',
          assignedToRSA: assigned,
          rsaAssignedAt: serverTimestamp(),
          status: 'processing_to_pfa',
          rsaReady: true,
          assignmentMode: 'skip_reviewer_routing',
          rsaAssignmentMode: 'round_robin_fallback',
          reviewerSkipped: true
        });
      } else {
        await markDirectRsaManualReview(subRef, 'round_robin_fallback_unavailable');
      }
    }
    console.error('Direct RSA round-robin transaction failed; using rotating fallback assignment.', {
      submissionId: subRef?.id || '',
      fallbackAssignedTo: assigned,
      rsaUsers,
      error
    });
  }
  return assigned;
}

// ==================== GET APPLICATION STAGE ====================
async function getApplicationStage(submission) {
  if (!submission) return 'Unknown';
  try {
    const status = String(submission.status || '').toLowerCase();
    if (status === 'cleared') return 'Cleared';
    if (status === 'paid') return 'Paid';
    if (status === 'sent_to_pfa' || status === 'rsa_submitted') return 'Sent to PFA';
    if (status === 'processing_to_pfa' || status === 'approved') return 'Processing to PFA';
    if (status === 'rejected') return 'Rejected - Fix Required';
    if (status === 'rejected_by_rsa') return 'Rejected by RSA - Fix Required';
    if (submission.assignedTo) {
      const reviewerName = await getUserFullName(submission.assignedTo);
      return `With Reviewer: ${reviewerName}`;
    }
    const historyQuery = query(
      collection(db, 'roundRobinAssignments'),
      where('submissionId', '==', submission.id),
      orderBy('assignedAt', 'desc')
    );
    const historySnap = await getDocs(historyQuery);
    if (!historySnap.empty) {
      const latest = historySnap.docs[0].data();
      if (latest.assignedTo) {
        const reviewerName = await getUserFullName(latest.assignedTo);
        return `With Reviewer: ${reviewerName}`;
      }
    }
    return 'Pending Assignment';
  } catch (error) {
    return 'Unknown';
  }
}

// ==================== GET USER FULL NAME BY EMAIL ====================
async function getUserFullName(email) {
  const normalized = normalizeEmail(email);
  if (!normalized) return 'Unknown';
  if (userFullNames.has(normalized)) return userFullNames.get(normalized);
  const fullName = await getUserFullNameShared(db, normalized);
  userFullNames.set(normalized, fullName);
  return fullName;
}

// ==================== DOM ELEMENTS ====================
const userName = document.getElementById('userName');
const userAvatar = document.getElementById('userAvatar');
const newUploadBtn = document.getElementById('newUploadBtn');
const bulkTemplateBtn = document.getElementById('bulkTemplateBtn');
const bulkImportBtn = document.getElementById('bulkImportBtn');
const bulkImportInput = document.getElementById('bulkImportInput');
const uploadModal = document.getElementById('uploadModal');
const editModal = document.getElementById('editModal');
const singleUploadModal = document.getElementById('singleUploadModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const closeEditModalBtn = document.getElementById('closeEditModalBtn');
const closeSingleModalBtn = document.getElementById('closeSingleModalBtn');
const cancelSingleBtn = document.getElementById('cancelSingleBtn');
const documentGrid = document.getElementById('documentGrid');
const optionalDocumentGrid = document.getElementById('optionalDocumentGrid');
const editDocumentGrid = document.getElementById('editDocumentGrid');
const customerNameInput = document.getElementById('customerName');
const submitCustomerBtn = document.getElementById('submitCustomerBtn');
const saveDraftBtn = document.getElementById('saveDraftBtn');
const submitEditBtn = document.getElementById('submitEditBtn');
const uploadedCountSpan = document.getElementById('uploadedCount');
const totalCountSpan = document.getElementById('totalCount');
const uploadProgressText = document.getElementById('uploadProgressText');
const notification = document.getElementById('notification');
const requiredDocumentsHeading = document.getElementById('requiredDocumentsHeading');
const batchUploadHint = document.getElementById('batchUploadHint');
const errorDetailModal = document.getElementById('errorDetailModal');
const errorDetailMessage = document.getElementById('errorDetailMessage');
const closeErrorDetailModalBtn = document.getElementById('closeErrorDetailModalBtn');
const dismissErrorDetailModalBtn = document.getElementById('dismissErrorDetailModalBtn');
const recentTableBody = document.getElementById('recentTableBody');
const editCustomerName = document.getElementById('editCustomerName');
const rejectionComment = document.getElementById('rejectionComment');
const singleUploadArea = document.getElementById('singleUploadArea');
const singleFileInput = document.getElementById('singleFileInput');
const singleFilePreview = document.getElementById('singleFilePreview');
const uploadModalTitle = document.getElementById('uploadModalTitle');
const uploadDocType = document.getElementById('uploadDocType');
const confirmSingleUpload = document.getElementById('confirmSingleUpload');
const pageTitle = document.getElementById('pageTitle');
const switchBackRoleLink = document.getElementById('switchBackRoleLink');
const switchBackRoleText = document.getElementById('switchBackRoleText');
const draftTableBody = document.getElementById('draftTableBody');
const pendingTableBody = document.getElementById('pendingTableBody');
const approvedTableBody = document.getElementById('approvedTableBody');
const rejectedTableBody = document.getElementById('rejectedTableBody');
const applicationsTableBody = document.getElementById('applicationsTableBody');
const applicationsTableHeadRow = applicationsTableBody?.closest('table')?.querySelector('thead tr');
const applicationsSearch = document.getElementById('applicationsSearch');
const uploaderRejectionReasonModal = document.getElementById('uploaderRejectionReasonModal');
const closeUploaderRejectionReasonModal = document.getElementById('closeUploaderRejectionReasonModal');
const closeUploaderRejectionReasonBtn = document.getElementById('closeUploaderRejectionReasonBtn');
const uploaderRejectionReasonCustomerName = document.getElementById('uploaderRejectionReasonCustomerName');
const uploaderRejectionReasonContact = document.getElementById('uploaderRejectionReasonContact');
const uploaderRejectionReasonHistory = document.getElementById('uploaderRejectionReasonHistory');
const activeCommissionTableBody = document.getElementById('activeCommissionTableBody');
const clearedCommissionTableBody = document.getElementById('clearedCommissionTableBody');
const paidCountBadge = document.getElementById('paidCount');
const activeCommissionCountEl = document.getElementById('activeCommissionCount');
const clearedCommissionCountEl = document.getElementById('clearedCommissionCount');
const paidTotal25El = document.getElementById('paidTotal25');
const paidTotal1El = document.getElementById('paidTotal1');
const clearedTotal25El = document.getElementById('clearedTotal25');
const clearedTotal1El = document.getElementById('clearedTotal1');
const activeTotal25Card = document.getElementById('activeTotal25Card');
const activeTotal1Card = document.getElementById('activeTotal1Card');
const clearedTotal25Card = document.getElementById('clearedTotal25Card');
const clearedTotal1Card = document.getElementById('clearedTotal1Card');
const activeCommissionTabBtn = document.getElementById('activeCommissionTabBtn');
const clearedCommissionTabBtn = document.getElementById('clearedCommissionTabBtn');
const activeCommissionSection = document.getElementById('activeCommissionSection');
const clearedCommissionSection = document.getElementById('clearedCommissionSection');
const activeCommissionSearch = document.getElementById('activeCommissionSearch');
const clearedCommissionSearch = document.getElementById('clearedCommissionSearch');
const activeCommissionStartDate = document.getElementById('activeCommissionStartDate');
const activeCommissionEndDate = document.getElementById('activeCommissionEndDate');
const clearedCommissionStartDate = document.getElementById('clearedCommissionStartDate');
const clearedCommissionEndDate = document.getElementById('clearedCommissionEndDate');
const viewerModal = document.getElementById('viewerModal');
const viewerFileName = document.getElementById('viewerFileName');
const documentViewer = document.getElementById('documentViewer');
const helpFab = document.getElementById('helpFab');
const sopHelpModal = document.getElementById('sopHelpModal');
const closeSopHelpModalBtn = document.getElementById('closeSopHelpModalBtn');
const closeSopHelpModalFooterBtn = document.getElementById('closeSopHelpModalFooterBtn');
const batchUploadBtn = document.getElementById('batchUploadBtn');
const batchFileInput = document.getElementById('batchFileInput');
const saveDetailsBtn = document.getElementById('saveDetailsBtn');
const resetDetailsBtn = document.getElementById('resetDetailsBtn');
const accountNoInput = document.getElementById('accountNo');
const accountBankSelect = document.getElementById('accountBank');
const accountNameInput = document.getElementById('accountName');
const accountLookupStatus = document.getElementById('accountLookupStatus');
const penNoInput = document.getElementById('penNo');
const penNoError = document.getElementById('penNoError');
const batchMappingModal = document.getElementById('batchMappingModal');
const batchMappingList = document.getElementById('batchMappingList');
const closeBatchMappingBtn = document.getElementById('closeBatchMappingBtn');
const cancelBatchMapping = document.getElementById('cancelBatchMapping');
const confirmBatchMapping = document.getElementById('confirmBatchMapping');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
const registeredAgentsTableBody = document.getElementById('registeredAgentsTableBody');
const customerAgentSelect = document.getElementById('customerAgent');
const agentAccountNumberInput = document.getElementById('agentAccountNumber');
const agentAccountBankSelect = document.getElementById('agentAccountBank');
const agentAccountNameInput = document.getElementById('agentAccountName');
const agentAccountLookupStatus = document.getElementById('agentAccountLookupStatus');
const pfaInput = document.getElementById('pfa');
const pfaOptionsList = document.getElementById('pfaOptionsList');
const uploadModalHeading = uploadModal?.querySelector('.modal-header h2');
const agentRegistrationForm = document.getElementById('agentRegistrationForm');
const agentRegistrationModal = document.getElementById('agentRegistrationModal');
const openAgentRegistrationBtn = document.getElementById('openAgentRegistrationBtn');
const closeAgentRegistrationModalBtn = document.getElementById('closeAgentRegistrationModalBtn');
const resetAgentFormBtn = document.getElementById('resetAgentFormBtn');
const submitAgentFormBtn = document.getElementById('submitAgentFormBtn');
let __batchFilesBuffer = [];
let currentCommissionTab = 'sent_to_pfa';
let currentUploaderApplicationTab = 'draft';
let currentUploaderPaidScope = 'mine';
let submissionInProgress = false;
const UPLOADER_DASHBOARD_TABS = ['overview', 'draft', 'applications', 'pending', 'approved', 'rejected', 'paid', 'register-agent', 'profile', 'help'];
const UPLOADER_APPLICATION_TABS = ['draft', 'pending', 'approved', 'rejected', 'sent_to_pfa', 'audit', 'paid', 'cleared'];
let registeredAgents = [];
let agentAccountLookupSequence = 0;
let agentAccountLookupDebounce = null;
let verifiedAgentAccountLookup = { accountNumber: '', bankCode: '', accountName: '' };

function getUploaderHashParts() {
  const hash = decodeURIComponent(String(window.location.hash || '').replace(/^#/, '')).trim();
  const [main = '', child = ''] = hash.split(':');
  return { main, child };
}

function getInitialUploaderTab() {
  const { main } = getUploaderHashParts();
  return UPLOADER_DASHBOARD_TABS.includes(main) ? main : 'overview';
}

function getInitialUploaderApplicationTab() {
  const { child } = getUploaderHashParts();
  return UPLOADER_APPLICATION_TABS.includes(child) ? child : currentUploaderApplicationTab;
}

function rememberUploaderTab(tabId, childTab = '') {
  if (!UPLOADER_DASHBOARD_TABS.includes(tabId)) return;
  const next = tabId === 'applications' && childTab ? `${tabId}:${childTab}` : tabId;
  if (window.location.hash === `#${next}`) return;
  history.replaceState(null, '', `#${next}`);
}

function renderProfileTab() {
  if (!profileNameEl && !profileEmailEl && !profileRoleEl && !profileStatusEl) return;
  const setProfileField = (el, value) => {
    if (!el) return;
    if ('value' in el) el.value = String(value ?? '');
    else el.textContent = String(value ?? '');
  };
  const fullName = currentUserProfile?.fullName || currentUser?.displayName || currentUser?.email || 'N/A';
  const registeredAt = currentUserProfile?.createdAt ? safeFormatDate(currentUserProfile.createdAt) : '-';
  const email = currentUserProfile?.email || currentUser?.email || 'N/A';
  const whatsapp = currentUserProfile?.whatsappNumber || currentUserProfile?.phone || '-';
  const location = currentUserProfile?.location || '-';
  const role = String(currentUserProfile?.role || 'uploader');
  const status = String(currentUserProfile?.status || 'active');
  setProfileField(profileNameEl, fullName);
  setProfileField(profileRegisteredAtEl, registeredAt);
  setProfileField(profileEmailEl, email);
  setProfileField(profileWhatsappEl, whatsapp);
  setProfileField(profileLocationEl, location);
  setProfileField(profileRoleEl, role.charAt(0).toUpperCase() + role.slice(1));
  setProfileField(profileStatusEl, status.charAt(0).toUpperCase() + status.slice(1));
  syncOriginatingTpField();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function showNewSubmissionConfirmModal({ customerName = '', agentPayload = {}, documentsCount = 0 } = {}) {
  return new Promise((resolve) => {
    const existingModal = document.getElementById('newSubmissionConfirmModal');
    if (existingModal) existingModal.remove();

    const hasAgent = Boolean(agentPayload?.agentName);
    const rows = [
      ['Customer', customerName || '-'],
      ['Agent', hasAgent ? agentPayload.agentName : 'No Agent'],
      ['Agent Account Number', hasAgent ? (agentPayload.agentAccountNumber || '-') : '-'],
      ['Documents Ready', String(documentsCount || 0)]
    ];

    const modal = document.createElement('div');
    modal.id = 'newSubmissionConfirmModal';
    modal.className = 'modal active';
    modal.style.zIndex = '4000';
    modal.innerHTML = `
      <div class="modal-content" style="max-width:520px;border-radius:14px;overflow:hidden;">
        <div class="modal-header" style="background:#003366;color:#fff;">
          <h2 style="color:#fff;margin:0;display:flex;align-items:center;gap:10px;font-size:20px;">
            <i class="fas fa-circle-check"></i> Confirm Submission
          </h2>
          <button class="close-btn" type="button" data-submit-confirm="cancel" aria-label="Close">&times;</button>
        </div>
        <div class="modal-body" style="padding:22px;">
          <p style="margin:0 0 16px;color:#475569;line-height:1.5;">
            Review these details before sending this application for processing.
          </p>
          <div style="display:grid;gap:10px;">
            ${rows.map(([label, value]) => `
              <div style="display:flex;justify-content:space-between;gap:16px;padding:12px 14px;border:1px solid #e2e8f0;border-radius:8px;background:#f8fafc;">
                <span style="color:#64748b;font-size:13px;font-weight:700;">${escapeHtml(label)}</span>
                <strong style="color:#0f172a;text-align:right;font-size:14px;">${escapeHtml(value)}</strong>
              </div>
            `).join('')}
          </div>
        </div>
        <div class="modal-footer" style="display:flex;justify-content:flex-end;gap:10px;padding:16px 22px;background:#f8fafc;border-top:1px solid #e2e8f0;">
          <button class="cancel-btn" type="button" data-submit-confirm="cancel">
            Cancel
          </button>
          <button class="submit-btn" type="button" data-submit-confirm="submit">
            <i class="fas fa-paper-plane"></i> Submit Application
          </button>
        </div>
      </div>
    `;

    let resolved = false;
    const close = (value) => {
      if (resolved) return;
      resolved = true;
      const submitBtn = modal.querySelector('[data-submit-confirm="submit"]');
      if (submitBtn) submitBtn.disabled = true;
      modal.remove();
      document.removeEventListener('keydown', onKeyDown);
      resolve(value);
    };
    const onKeyDown = (event) => {
      if (event.key === 'Escape') close(false);
    };

    modal.addEventListener('click', (event) => {
      if (event.target === modal) close(false);
      const action = event.target?.closest?.('[data-submit-confirm]')?.dataset?.submitConfirm;
      if (action === 'cancel') close(false);
      if (action === 'submit') close(true);
    });
    document.addEventListener('keydown', onKeyDown);
    (document.body || document.documentElement).appendChild(modal);
  });
}

function getUploadModalMode() {
  if (currentEditId) return 'fix';
  if (currentDraftId) return 'draft';
  return 'new';
}

function setUploadModalHeading(mode = getUploadModalMode()) {
  if (!uploadModalHeading) return;
  if (mode === 'draft') {
    uploadModalHeading.innerHTML = '<i class="fas fa-file-pen"></i> Resume Draft Submission';
    return;
  }
  if (mode === 'fix') {
    uploadModalHeading.innerHTML = '<i class="fas fa-edit"></i> Correct & Re-upload';
    return;
  }
  uploadModalHeading.innerHTML = '<i class="fas fa-user-plus"></i> New Customer Submission';
}

function syncUploadModalControlState(mode = getUploadModalMode()) {
  if (saveDraftBtn) {
    saveDraftBtn.style.display = mode === 'fix' ? 'none' : 'inline-flex';
  }
}

function getFormValue(id) {
  return document.getElementById(id)?.value ?? '';
}

function normalizeCustomerPhone(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeDigitsOnly(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeCustomerNameKey(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}

function getSubmissionCustomerEmail(submission) {
  return normalizeEmail(submission?.customerDetails?.email || submission?.customerEmail || submission?.email || '');
}

function getSubmissionCustomerPhone(submission) {
  return normalizeCustomerPhone(submission?.customerDetails?.phone || submission?.customerPhone || submission?.phone || '');
}

function getSubmissionCustomerAccountNo(submission) {
  return normalizeDigitsOnly(submission?.customerDetails?.accountNo || submission?.accountNo || '');
}

function getSubmissionCustomerNin(submission) {
  return normalizeDigitsOnly(submission?.customerDetails?.nin || submission?.customerNIN || submission?.nin || '');
}

function normalizePenNumber(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function getSubmissionCustomerPenNo(submission) {
  return normalizePenNumber(submission?.customerDetails?.penNoNormalized || submission?.penNoNormalized || submission?.customerDetails?.penNo || submission?.penNo || '');
}

function getSubmissionCustomerNameKey(submission) {
  return normalizeCustomerNameKey(submission?.customerName || submission?.customerDetails?.name || '');
}

function submissionUniqueKeyDocId(type, value) {
  return encodeURIComponent(`${String(type || '').trim().toLowerCase()}:${String(value || '').trim()}`);
}

function buildSubmissionUniqueKeysFromDetails(customerDetails = {}, customerName = '') {
  const keys = [
    { type: 'account_number', label: 'account number', value: normalizeDigitsOnly(customerDetails.accountNo) },
    { type: 'nin', label: 'NIN', value: normalizeDigitsOnly(customerDetails.nin) },
    { type: 'pen', label: 'PEN', value: normalizePenNumber(customerDetails.penNoNormalized || customerDetails.penNo) },
    { type: 'phone', label: 'phone number', value: normalizeCustomerPhone(customerDetails.phone) },
    { type: 'customer_name', label: 'customer name', value: normalizeCustomerNameKey(customerName || customerDetails.name) }
  ];
  return keys
    .filter((item) => item.value)
    .filter((item) => {
      if (item.type === 'account_number') return item.value.length === 10;
      if (item.type === 'nin') return item.value.length === 11;
      if (item.type === 'phone') return item.value.length >= 10;
      if (item.type === 'customer_name') return item.value.length >= 3;
      return true;
    });
}

function getSubmissionUniqueKeyConflicts(existingDoc, submissionId = '') {
  if (!existingDoc?.exists()) return null;
  const data = existingDoc.data() || {};
  const existingSubmissionId = String(data.submissionId || '').trim();
  if (existingSubmissionId && existingSubmissionId === String(submissionId || '').trim()) return null;
  return data;
}

function formatSubmissionDuplicateLockMessage(conflicts = []) {
  const labels = [...new Set(conflicts.map((item) => item.label).filter(Boolean))];
  const first = conflicts[0]?.data || {};
  const customerName = String(first.customerName || '').trim();
  const suffix = customerName ? `\nExisting customer: ${customerName}` : '';
  return `Duplicate application blocked. Existing ${labels.join(', ')} already found in the database.${suffix}`;
}

async function saveSubmissionWithUniqueLocks(subRef, submissionPayload = {}, { mode = 'set' } = {}) {
  const submissionId = subRef?.id || '';
  const keys = buildSubmissionUniqueKeysFromDetails(submissionPayload.customerDetails || {}, submissionPayload.customerName || '');
  await runTransaction(db, async (tx) => {
    const keyRefs = keys.map((item) => ({
      ...item,
      ref: doc(db, 'submissionUniqueKeys', submissionUniqueKeyDocId(item.type, item.value))
    }));
    const keySnaps = await Promise.all(keyRefs.map((item) => tx.get(item.ref)));
    const conflicts = keySnaps
      .map((snap, index) => ({ ...keyRefs[index], data: getSubmissionUniqueKeyConflicts(snap, submissionId) }))
      .filter((item) => item.data);
    if (conflicts.length) {
      const error = new Error(formatSubmissionDuplicateLockMessage(conflicts));
      error.code = 'duplicate-submission-lock';
      throw error;
    }

    if (mode === 'update') {
      tx.update(subRef, submissionPayload);
    } else {
      tx.set(subRef, submissionPayload, { merge: mode === 'merge' });
    }

    keyRefs.forEach((item) => {
      tx.set(item.ref, {
        type: item.type,
        label: item.label,
        value: item.value,
        submissionId,
        customerName: submissionPayload.customerName || submissionPayload.customerDetails?.name || '',
        uploadedBy: normalizeEmail(submissionPayload.uploadedBy || currentUser?.email || ''),
        updatedAt: serverTimestamp()
      }, { merge: true });
    });
  });
}

function getPenNumberQueryVariants(penNo = '') {
  const raw = String(penNo || '').trim();
  const normalized = normalizePenNumber(raw);
  return Array.from(new Set([
    raw,
    raw.toUpperCase(),
    raw.toLowerCase(),
    normalized,
    normalized.toUpperCase()
  ].filter(Boolean))).slice(0, 10);
}

function isCurrentEditableSubmission(submissionId) {
  return Boolean(submissionId) && (submissionId === currentDraftId || submissionId === currentEditId);
}

function getEditableSubmissionById(submissionId = '') {
  if (!submissionId) return null;
  return allSubmissions.find((submission) => submission.id === submissionId) || null;
}

function getCurrentEditableSubmission() {
  return getEditableSubmissionById(currentEditId || currentDraftId || '');
}

function isUnchangedEditablePenNumber(penNo = '', submissionId = currentEditId || currentDraftId || '') {
  const normalizedPenNo = normalizePenNumber(penNo);
  if (!normalizedPenNo || !submissionId) return false;
  const submission = getEditableSubmissionById(submissionId);
  return Boolean(submission && getSubmissionCustomerPenNo(submission) === normalizedPenNo);
}

async function findExistingPenNumberSubmissions({ penNo = '', excludeSubmissionId = '' } = {}) {
  const normalizedPenNo = normalizePenNumber(penNo);
  if (!normalizedPenNo) return [];

  const variants = getPenNumberQueryVariants(penNo);
  const matches = new Map();
  const addSnapshotMatches = (snapshot) => {
    snapshot.forEach((docSnap) => {
      if (excludeSubmissionId && docSnap.id === excludeSubmissionId) return;
      const submission = { id: docSnap.id, ...(docSnap.data() || {}) };
      if (getSubmissionCustomerPenNo(submission) === normalizedPenNo) {
        matches.set(docSnap.id, submission);
      }
    });
  };

  const queries = [
    query(collection(db, 'submissions'), where('customerDetails.penNo', 'in', variants)),
    query(collection(db, 'submissions'), where('penNo', 'in', variants)),
    query(collection(db, 'submissions'), where('customerDetails.penNoNormalized', '==', normalizedPenNo)),
    query(collection(db, 'submissions'), where('penNoNormalized', '==', normalizedPenNo))
  ];

  try {
    const snapshots = await Promise.all(queries.map((q) => getDocs(q).catch(() => null)));
    snapshots.filter(Boolean).forEach(addSnapshotMatches);
  } catch (_) {}

  return Array.from(matches.values());
}

async function findExistingSubmissionFieldMatches({ phone = '', accountNo = '', nin = '', customerName = '', excludeSubmissionId = '' } = {}) {
  const specs = [
    { value: normalizeCustomerPhone(phone), compare: getSubmissionCustomerPhone, fields: ['customerDetails.phone', 'customerPhone', 'phone'] },
    { value: normalizeDigitsOnly(accountNo), compare: getSubmissionCustomerAccountNo, fields: ['customerDetails.accountNo', 'accountNo'] },
    { value: normalizeDigitsOnly(nin), compare: getSubmissionCustomerNin, fields: ['customerDetails.nin', 'customerNIN', 'nin'] },
    { value: String(customerName || '').trim(), compare: (sub) => String(sub?.customerName || sub?.customerDetails?.name || '').trim(), fields: ['customerName', 'customerDetails.name'] }
  ].filter((spec) => spec.value);
  const matches = new Map();
  await Promise.all(specs.flatMap((spec) => spec.fields.map(async (field) => {
    try {
      const snap = await getDocs(query(collection(db, 'submissions'), where(field, '==', spec.value)));
      snap.forEach((docSnap) => {
        if (excludeSubmissionId && docSnap.id === excludeSubmissionId) return;
        const submission = { id: docSnap.id, ...(docSnap.data() || {}) };
        const left = field.includes('name') || field === 'customerName'
          ? normalizeCustomerNameKey(spec.compare(submission))
          : String(spec.compare(submission) || '');
        const right = field.includes('name') || field === 'customerName'
          ? normalizeCustomerNameKey(spec.value)
          : String(spec.value || '');
        if (left && left === right) matches.set(docSnap.id, submission);
      });
    } catch (_) {}
  })));
  return Array.from(matches.values());
}

async function findSubmissionUniqueLockConflicts(customerDetails = {}, customerName = '', excludeSubmissionId = '') {
  const keys = buildSubmissionUniqueKeysFromDetails(customerDetails, customerName);
  if (!keys.length) return [];
  const docs = await Promise.all(keys.map(async (item) => {
    const snap = await getDoc(doc(db, 'submissionUniqueKeys', submissionUniqueKeyDocId(item.type, item.value))).catch(() => null);
    const data = getSubmissionUniqueKeyConflicts(snap, excludeSubmissionId);
    return data ? { ...item, data } : null;
  }));
  return docs.filter(Boolean);
}

function findDuplicateCustomerContacts({ email = '', phone = '', accountNo = '', nin = '', penNo = '', customerName = '', excludeSubmissionId = '' } = {}) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedPhone = normalizeCustomerPhone(phone);
  const normalizedAccountNo = normalizeDigitsOnly(accountNo);
  const normalizedNin = normalizeDigitsOnly(nin);
  const normalizedPenNo = normalizePenNumber(penNo);
  const normalizedCustomerName = normalizeCustomerNameKey(customerName);
  return allSubmissions.filter((submission) => {
    if (excludeSubmissionId && submission.id === excludeSubmissionId) return false;
    if (isCurrentEditableSubmission(submission.id) && submission.id === excludeSubmissionId) return false;
    const emailMatch = normalizedEmail && getSubmissionCustomerEmail(submission) === normalizedEmail;
    const phoneMatch = normalizedPhone && getSubmissionCustomerPhone(submission) === normalizedPhone;
    const accountMatch = normalizedAccountNo && getSubmissionCustomerAccountNo(submission) === normalizedAccountNo;
    const ninMatch = normalizedNin && getSubmissionCustomerNin(submission) === normalizedNin;
    const penMatch = normalizedPenNo && getSubmissionCustomerPenNo(submission) === normalizedPenNo;
    const nameMatch = normalizedCustomerName && getSubmissionCustomerNameKey(submission) === normalizedCustomerName;
    return emailMatch || phoneMatch || accountMatch || ninMatch || penMatch || nameMatch;
  });
}

function formatDuplicateCustomerSummary(submission, { email = '', phone = '', accountNo = '', nin = '', penNo = '', customerName = '' } = {}) {
  const emailMatch = normalizeEmail(email) && getSubmissionCustomerEmail(submission) === normalizeEmail(email);
  const phoneMatch = normalizeCustomerPhone(phone) && getSubmissionCustomerPhone(submission) === normalizeCustomerPhone(phone);
  const accountMatch = normalizeDigitsOnly(accountNo) && getSubmissionCustomerAccountNo(submission) === normalizeDigitsOnly(accountNo);
  const ninMatch = normalizeDigitsOnly(nin) && getSubmissionCustomerNin(submission) === normalizeDigitsOnly(nin);
  const penMatch = normalizePenNumber(penNo) && getSubmissionCustomerPenNo(submission) === normalizePenNumber(penNo);
  const nameMatch = normalizeCustomerNameKey(customerName) && getSubmissionCustomerNameKey(submission) === normalizeCustomerNameKey(customerName);
  const reasons = [];
  if (emailMatch) reasons.push('email');
  if (phoneMatch) reasons.push('phone');
  if (accountMatch) reasons.push('account number');
  if (ninMatch) reasons.push('NIN');
  if (penMatch) reasons.push('PEN number');
  if (nameMatch) reasons.push('customer name');
  const duplicateCustomerName = String(submission?.customerName || submission?.customerDetails?.name || 'Unknown Customer').trim() || 'Unknown Customer';
  const status = String(submission?.status || 'the system').replace(/_/g, ' ');
  return `- ${duplicateCustomerName} (${reasons.join(' and ')}) [${status}]`;
}

async function validateCustomerDuplicateContact({ email = '', phone = '', accountNo = '', nin = '', penNo = '', customerName = '', excludeSubmissionId = '' } = {}) {
  const localMatches = findDuplicateCustomerContacts({ email, phone, accountNo, nin, penNo, customerName, excludeSubmissionId });
  const remotePenMatches = await findExistingPenNumberSubmissions({ penNo, excludeSubmissionId });
  const remoteFieldMatches = await findExistingSubmissionFieldMatches({ phone, accountNo, nin, customerName, excludeSubmissionId });
  const lockConflicts = await findSubmissionUniqueLockConflicts({ phone, accountNo, nin, penNo }, customerName, excludeSubmissionId);
  const existingMatches = Array.from(new Map([...localMatches, ...remotePenMatches, ...remoteFieldMatches].map((submission) => [submission.id, submission])).values());
  if (!existingMatches.length && !lockConflicts.length) return null;

  const reasons = [];
  if (normalizeEmail(email) && existingMatches.some((submission) => getSubmissionCustomerEmail(submission) === normalizeEmail(email))) reasons.push('email address');
  if ((normalizeCustomerPhone(phone) && existingMatches.some((submission) => getSubmissionCustomerPhone(submission) === normalizeCustomerPhone(phone))) || lockConflicts.some((item) => item.type === 'phone')) reasons.push('phone number');
  if ((normalizeDigitsOnly(accountNo) && existingMatches.some((submission) => getSubmissionCustomerAccountNo(submission) === normalizeDigitsOnly(accountNo))) || lockConflicts.some((item) => item.type === 'account_number')) reasons.push('account number');
  if ((normalizeDigitsOnly(nin) && existingMatches.some((submission) => getSubmissionCustomerNin(submission) === normalizeDigitsOnly(nin))) || lockConflicts.some((item) => item.type === 'nin')) reasons.push('NIN');
  if ((normalizePenNumber(penNo) && existingMatches.some((submission) => getSubmissionCustomerPenNo(submission) === normalizePenNumber(penNo))) || lockConflicts.some((item) => item.type === 'pen')) reasons.push('PEN number');
  if ((normalizeCustomerNameKey(customerName) && existingMatches.some((submission) => getSubmissionCustomerNameKey(submission) === normalizeCustomerNameKey(customerName))) || lockConflicts.some((item) => item.type === 'customer_name')) reasons.push('customer name');
  const duplicateLines = existingMatches.map((submission) => formatDuplicateCustomerSummary(submission, { email, phone, accountNo, nin, penNo, customerName }));
  lockConflicts.forEach((item) => {
    const lockedName = String(item.data?.customerName || 'Existing application').trim();
    duplicateLines.push(`- ${lockedName} (${item.label}) [existing application]`);
  });
  return {
    submissions: existingMatches,
    message: `Duplicate application blocked. Existing ${[...new Set(reasons)].join(' and ')} already found in the database.\n\nMatching record(s):\n${duplicateLines.join('\n')}`
  };
}

function clearPenNoError() {
  if (penNoError) {
    penNoError.textContent = '';
    penNoError.style.display = 'none';
  }
  if (penNoInput) {
    penNoInput.style.borderColor = '';
    penNoInput.style.boxShadow = '';
    penNoInput.removeAttribute('aria-invalid');
  }
}

function showPenNoError(message = 'PEN No already existed.') {
  if (penNoError) {
    penNoError.textContent = message;
    penNoError.style.display = 'block';
  }
  if (penNoInput) {
    penNoInput.style.borderColor = '#dc2626';
    penNoInput.style.boxShadow = '0 0 0 3px rgba(220, 38, 38, 0.12)';
    penNoInput.setAttribute('aria-invalid', 'true');
  }
}

function getDuplicateCustomerDisplayName(duplicate) {
  const submission = duplicate?.submissions?.[0] || {};
  return String(submission?.customerName || submission?.customerDetails?.name || '').trim();
}

function formatPenNoDuplicateMessage(duplicate) {
  const customerName = getDuplicateCustomerDisplayName(duplicate);
  return customerName ? `PEN No already existed.\n"${customerName}"` : 'PEN No already existed.';
}

async function validatePenNumberAvailable(penNo = '', excludeSubmissionId = '', options = {}) {
  const { inline = false, notify = true } = options || {};
  if (isUnchangedEditablePenNumber(penNo, excludeSubmissionId)) {
    if (inline) clearPenNoError();
    return true;
  }
  const duplicate = await validateCustomerDuplicateContact({ penNo, excludeSubmissionId });
  if (!duplicate) {
    if (inline) clearPenNoError();
    return true;
  }
  if (inline) {
    showPenNoError(formatPenNoDuplicateMessage(duplicate));
  } else if (notify) {
    showNotification(duplicate.message, 'error');
  }
  return false;
}

async function validatePenNoFieldInline() {
  const currentSequence = ++penNoValidationSequence;
  const penNo = penNoInput?.value?.trim() || '';
  if (!penNo) {
    clearPenNoError();
    return true;
  }
  const excludeSubmissionId = currentEditId || currentDraftId || '';
  if (isUnchangedEditablePenNumber(penNo, excludeSubmissionId)) {
    if (currentSequence !== penNoValidationSequence) return true;
    clearPenNoError();
    return true;
  }
  const duplicate = await validateCustomerDuplicateContact({ penNo, excludeSubmissionId });
  if (currentSequence !== penNoValidationSequence) return true;
  if (duplicate) {
    showPenNoError(formatPenNoDuplicateMessage(duplicate));
    return false;
  }
  clearPenNoError();
  return true;
}

function collectDraftableCustomerDetails() {
  const rsaBalance = parseMoney(getFormValue('rsaBalance'));
  const derivedProperty = rsaBalance ? determinePropertyByRsa(rsaBalance) : null;
  return {
    name: String(getFormValue('customerName')).trim(),
    dob: getFormValue('customerDob'),
    email: String(getFormValue('customerEmail')).trim(),
    phone: String(getFormValue('customerPhone')).trim(),
    nin: String(getFormValue('customerNIN')).trim(),
    address: String(getFormValue('customerAddress')).trim(),
    accountNo: String(getFormValue('accountNo')).trim(),
    accountBank: getSelectedAccountBank().name,
    accountBankCode: getSelectedAccountBank().code,
    accountName: String(verifiedAccountLookup.accountName || getFormValue('customerName')).trim(),
    employer: String(getFormValue('employer')).trim(),
    originatingTP: String(getFormValue('originatingTP')).trim() || String(currentUserProfile?.location || '').trim(),
    mortgageLoanApplicationFormDate: getFormValue('mortgageLoanApplicationFormDate'),
    pfa: String(getFormValue('pfa')).trim(),
    penNo: String(getFormValue('penNo')).trim(),
    penNoNormalized: normalizePenNumber(getFormValue('penNo')),
    rsaStatementDate: getFormValue('rsaStatementDate'),
    rsaBalance: getFormValue('rsaBalance'),
    rsa25: getFormValue('rsa25Percent'),
    propertyType: String(getFormValue('propertyType')).trim() || String(derivedProperty?.name || '').trim(),
    houseNumber: String(getFormValue('houseNumber')).trim(),
    tenor: getFormValue('tenor'),
    propertyValue: getFormValue('propertyValue') || (derivedProperty ? String(derivedProperty.value) : ''),
    facilityFee: getFormValue('facilityFee') || (derivedProperty ? String(derivedProperty.fee) : ''),
    loanAmount: normalizeLoanAmountValue(getFormValue('loanAmount') || '')
  };
}

function collectSubmissionDocuments() {
  const documents = [];
  Object.entries(currentCustomerUploads || {}).forEach(([type, files]) => {
    const latestFile = getLatestUploadedDoc(files || []);
    if (latestFile) documents.push({ documentType: type, ...latestFile });
  });
  return documents;
}

function getUploadedDocTimestamp(docItem = {}) {
  return Math.max(
    getStageTimestampMillis(docItem?.uploadedAt),
    Number(docItem?.localAddedAt || 0)
  );
}

function getLatestUploadedDoc(files = []) {
  const validFiles = (Array.isArray(files) ? files : [])
    .filter((file) => file && (file.fileUrl || file.fileId || file.name));
  if (!validFiles.length) return null;
  return validFiles.reduce((latest, file) => {
    if (!latest) return file;
    const fileTime = getUploadedDocTimestamp(file);
    const latestTime = getUploadedDocTimestamp(latest);
    return fileTime >= latestTime ? file : latest;
  }, null);
}

function getUploadedDocumentTypes() {
  return Object.entries(currentCustomerUploads || {})
    .filter(([, files]) => Boolean(getLatestUploadedDoc(files || [])))
    .map(([type]) => type);
}

function getEffectiveSubmissionDocuments(submission = null) {
  const rawDocs = Array.isArray(submission?.documents) ? submission.documents : [];
  if (rawDocs.length <= 1) return rawDocs;

  const latestByType = new Map();
  rawDocs.forEach((docItem, index) => {
    const type = String(docItem?.documentType || '').trim();
    const key = type || `__index_${index}`;
    const previous = latestByType.get(key);
    if (!previous || getUploadedDocTimestamp(docItem) >= getUploadedDocTimestamp(previous)) {
      latestByType.set(key, docItem);
    }
  });

  const orderedTypes = Array.isArray(submission?.documentTypes) ? submission.documentTypes : [];
  const orderedDocs = orderedTypes
    .map((type) => latestByType.get(String(type || '').trim()))
    .filter(Boolean);
  const included = new Set(orderedDocs);
  const remainder = Array.from(latestByType.values()).filter((docItem) => !included.has(docItem));
  return [...orderedDocs, ...remainder];
}

function getStoredAgentSelectionValue(submission = {}) {
  const storedAgentId = String(submission?.agentId || '').trim();
  return storedAgentId || '';
}

function getKnownAgentById(agentId = '') {
  const normalizedId = String(agentId || '').trim();
  if (!normalizedId) return null;
  return approvedAgents.find((agent) => String(agent?.id || '').trim() === normalizedId)
    || registeredAgents.find((agent) => String(agent?.id || '').trim() === normalizedId)
    || null;
}

function buildAgentSnapshotFromRecord(agent = null) {
  if (!agent) {
    return {
      agentId: '',
      agentName: '',
      agentContactNumber: '',
      agentAccountNumber: '',
      agentAccountBank: ''
    };
  }
  return {
    agentId: String(agent?.id || agent?.agentId || '').trim(),
    agentName: String(agent?.fullName || agent?.agentName || '').trim(),
    agentContactNumber: String(agent?.contactNumber || agent?.agentContactNumber || '').trim(),
    agentAccountNumber: String(agent?.accountNumber || agent?.agentAccountNumber || '').trim(),
    agentAccountBank: String(agent?.accountBank || agent?.agentAccountBank || '').trim()
  };
}

function getSubmissionAgentSnapshot(submission = {}) {
  const agentId = String(submission?.agentId || '').trim();
  if (!agentId) return buildAgentSnapshotFromRecord(null);
  const linkedAgent = getKnownAgentById(agentId);
  if (linkedAgent) return buildAgentSnapshotFromRecord(linkedAgent);
  return {
    agentId,
    agentName: String(submission?.agentName || '').trim(),
    agentContactNumber: String(submission?.agentContactNumber || '').trim(),
    agentAccountNumber: String(submission?.agentAccountNumber || '').trim(),
    agentAccountBank: String(submission?.agentAccountBank || '').trim()
  };
}

function normalizeSubmissionAgentFields(submission = {}) {
  const snapshot = getSubmissionAgentSnapshot(submission);
  return {
    ...submission,
    agentId: snapshot.agentId,
    agentName: snapshot.agentName,
    agentContactNumber: snapshot.agentContactNumber,
    agentAccountNumber: snapshot.agentAccountNumber,
    agentAccountBank: snapshot.agentAccountBank
  };
}

function getSubmissionAgentDisplayName(submission = {}) {
  return String(getSubmissionAgentSnapshot(submission).agentName || '').trim() || 'No Agent';
}

function syncCurrentSubmissionAgentFallbackFromSubmission(submission = {}) {
  const snapshot = getSubmissionAgentSnapshot(submission);
  currentSubmissionAgentFallback = snapshot.agentId
    ? {
        value: snapshot.agentId,
        id: snapshot.agentId,
        fullName: snapshot.agentName,
        contactNumber: snapshot.agentContactNumber,
        accountNumber: snapshot.agentAccountNumber,
        accountBank: snapshot.agentAccountBank
      }
    : null;
}

function getSelectedAgentPayload() {
  const selectedAgentId = String(customerAgentSelect?.value || '').trim();
  const selectedAgent = approvedAgents.find((agent) => agent.id === selectedAgentId) || null;
  if (selectedAgent) {
    return {
      selectedAgentId,
      selectedAgent,
      agentId: selectedAgent.id || '',
      agentName: selectedAgent.fullName || '',
      agentContactNumber: selectedAgent.contactNumber || '',
      agentAccountNumber: selectedAgent.accountNumber || '',
      agentAccountBank: selectedAgent.accountBank || ''
    };
  }
  if (currentSubmissionAgentFallback && selectedAgentId === currentSubmissionAgentFallback.value) {
    return {
      selectedAgentId,
      selectedAgent: null,
      agentId: currentSubmissionAgentFallback.id || '',
      agentName: currentSubmissionAgentFallback.fullName || '',
      agentContactNumber: currentSubmissionAgentFallback.contactNumber || '',
      agentAccountNumber: currentSubmissionAgentFallback.accountNumber || '',
      agentAccountBank: currentSubmissionAgentFallback.accountBank || ''
    };
  }
  return {
    selectedAgentId,
    selectedAgent: null,
    agentId: '',
    agentName: '',
    agentContactNumber: '',
    agentAccountNumber: '',
    agentAccountBank: ''
  };
}

function setUploadedDocForType(docType, docEntry) {
  if (!docType || !docEntry) return;
  currentCustomerUploads[docType] = [docEntry];
}

async function ensureStoredHouseNumber(customerDetails, options = {}) {
  const details = customerDetails || {};
  const propertyType = String(details.propertyType || '').trim();
  if (!propertyType) return details;

  const existingHouseNumber = String(details.houseNumber || '').trim();
  if (existingHouseNumber && currentHouseNumberReserved) {
    return {
      ...details,
      houseNumber: existingHouseNumber
    };
  }

  const reservedHouseNumber = await reserveHouseNumber(propertyType);
  if (!reservedHouseNumber) return details;
  currentHouseNumberReserved = true;
  const houseNumberEl = document.getElementById('houseNumber');
  if (!options.skipDomUpdate && houseNumberEl) {
    houseNumberEl.value = reservedHouseNumber;
  }
  return {
    ...details,
    houseNumber: reservedHouseNumber
  };
}

function hasAnyDraftContent() {
  const details = collectDraftableCustomerDetails();
  return Object.values(details).some((value) => String(value || '').trim()) || collectSubmissionDocuments().length > 0;
}

function hydrateCurrentUploads(documents = []) {
  const existingUploads = {};
  (documents || []).forEach((item) => {
    if (!item?.documentType) return;
    const type = item.documentType;
    const previous = existingUploads[type]?.[0] || null;
    if (!previous || getUploadedDocTimestamp(item) >= getUploadedDocTimestamp(previous)) {
      existingUploads[type] = [item];
    }
  });
  currentCustomerUploads = existingUploads;
}

function applyDraftFormValues(sub) {
  const details = sub?.customerDetails || {};
  const pick = (keys, fallback = '') => {
    for (const key of keys) {
      const detailValue = details?.[key];
      if (detailValue !== undefined && detailValue !== null && String(detailValue).trim() !== '') return String(detailValue);
      const rootValue = sub?.[key];
      if (rootValue !== undefined && rootValue !== null && String(rootValue).trim() !== '') return String(rootValue);
    }
    return fallback;
  };

  const fieldMap = [
    { id: 'customerName', keys: ['name', 'customerName'], fallback: sub?.customerName || '' },
    { id: 'customerDob', keys: ['dob', 'dateOfBirth', 'customerDob'] },
    { id: 'customerEmail', keys: ['email', 'customerEmail'] },
    { id: 'customerPhone', keys: ['phone', 'customerPhone'] },
    { id: 'customerNIN', keys: ['nin', 'customerNIN'] },
    { id: 'customerAddress', keys: ['address', 'customerAddress'] },
    { id: 'accountNo', keys: ['accountNo'] },
    { id: 'accountBank', keys: ['accountBankCode', 'bankCode'], fallback: DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE },
    { id: 'employer', keys: ['employer'] },
    { id: 'originatingTP', keys: ['originatingTP'], fallback: String(currentUserProfile?.location || '').trim() },
    { id: 'mortgageLoanApplicationFormDate', keys: ['mortgageLoanApplicationFormDate'] },
    { id: 'pfa', keys: ['pfa', 'pfaName'] },
    { id: 'penNo', keys: ['penNo'] },
    { id: 'rsaStatementDate', keys: ['rsaStatementDate'] },
    { id: 'rsaBalance', keys: ['rsaBalance'] },
    { id: 'propertyType', keys: ['propertyType'] },
    { id: 'houseNumber', keys: ['houseNumber'], fallback: sub?.houseNumber || '' },
    { id: 'propertyValue', keys: ['propertyValue'] },
    { id: 'facilityFee', keys: ['facilityFee'] },
    { id: 'loanAmount', keys: ['loanAmount'] },
    { id: 'tenor', keys: ['tenor'] }
  ];

  fieldMap.forEach((field) => {
    const el = document.getElementById(field.id);
    if (!el) return;
    const rawValue = pick(field.keys, field.fallback || '');
    if (field.id === 'loanAmount') {
      el.value = normalizeLoanAmountValue(rawValue);
    } else if (field.id === 'accountBank') {
      if (el.tagName !== 'SELECT') {
        el.value = DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME;
        return;
      }
      const bankName = pick(['accountBank', 'bankName'], '');
      const bank = accountLookupBanks.find((item) => item.code === rawValue || item.name === bankName);
      el.value = bank?.code || rawValue;
    } else {
      el.value = rawValue;
    }
  });
  if (accountNoInput && accountBankSelect?.value && customerNameInput?.value) {
    verifiedAccountLookup = {
      accountNumber: String(accountNoInput.value || '').replace(/\D/g, ''),
      bankCode: getSelectedAccountBank().code,
      accountName: String(customerNameInput.value || '').trim()
    };
    setAccountLookupStatus(`Verified: ${verifiedAccountLookup.accountName}`, 'success');
  }

  const rsa25Value = pick(['rsa25', 'rsa25Percent'], '');
  const rsa25PercentEl = document.getElementById('rsa25Percent');
  if (rsa25PercentEl) rsa25PercentEl.value = rsa25Value;
  const rsa25FormattedEl = document.getElementById('rsa25Formatted');
  if (rsa25FormattedEl) rsa25FormattedEl.textContent = rsa25Value ? formatCurrency(rsa25Value) : '';
  const propertyValueFormattedEl = document.getElementById('propertyValueFormatted');
  if (propertyValueFormattedEl) propertyValueFormattedEl.textContent = getFormValue('propertyValue') ? formatCurrency(getFormValue('propertyValue')) : '';
  const facilityFeeFormattedEl = document.getElementById('facilityFeeFormatted');
  if (facilityFeeFormattedEl) facilityFeeFormattedEl.textContent = getFormValue('facilityFee') ? formatCurrency(getFormValue('facilityFee')) : '';
  const loanAmountFormattedEl = document.getElementById('loanAmountFormatted');
  if (loanAmountFormattedEl) loanAmountFormattedEl.textContent = getFormValue('loanAmount') ? formatCurrency(getFormValue('loanAmount')) : '';
}

function enableDraftEditingState() {
  customerDetailsSaved = true;
  if (batchUploadBtn) batchUploadBtn.disabled = false;
  if (documentGrid) {
    documentGrid.style.pointerEvents = 'auto';
    documentGrid.style.opacity = '1';
  }
  if (optionalDocumentGrid) {
    optionalDocumentGrid.style.pointerEvents = 'auto';
    optionalDocumentGrid.style.opacity = '1';
  }
}

async function persistCurrentDraft({ silent = false, source = 'manual' } = {}) {
  if (!assertWritable('Draft saving')) return null;
  const uploaderEmail = normalizeEmail(currentUser?.email);
  if (!uploaderEmail) throw new Error('Unable to determine uploader email for draft save.');
  if (!hasAnyDraftContent()) {
    if (!silent) showNotification('Add customer details or documents before saving draft.', 'error');
    return null;
  }

  const customerDetails = collectDraftableCustomerDetails();
  if (!String(customerDetails.name || '').trim()) {
    if (!silent) showNotification('Customer name is required before saving a draft.', 'error');
    return null;
  }
  const storedCustomerDetails = await ensureStoredHouseNumber(customerDetails);
  const documents = collectSubmissionDocuments();
  const documentTypes = getUploadedDocumentTypes();
  const agentPayload = getSelectedAgentPayload();
  const payload = {
    customerName: storedCustomerDetails.name || 'Untitled Draft',
    customerDetails: storedCustomerDetails,
    uploadedBy: uploaderEmail,
    status: 'draft',
    comment: '',
    documents,
    documentTypes,
    houseNumber: storedCustomerDetails.houseNumber || '',
    penNoNormalized: normalizePenNumber(storedCustomerDetails.penNo),
    agentId: agentPayload.agentId,
    agentName: agentPayload.agentName,
    agentContactNumber: agentPayload.agentContactNumber,
    agentAccountNumber: agentPayload.agentAccountNumber,
    agentAccountBank: agentPayload.agentAccountBank,
    draftSource: source,
    draftSavedAt: serverTimestamp()
  };

  if (currentDraftId) {
    await updateDoc(doc(db, 'submissions', currentDraftId), payload);
  } else {
    const created = await addDoc(collection(db, 'submissions'), {
      ...payload,
      uploadedAt: serverTimestamp()
    });
    currentDraftId = created.id;
  }

  enableDraftEditingState();
  setUploadModalHeading('draft');
  syncUploadModalControlState('draft');
  if (!silent) showNotification('Draft saved successfully.', 'success');
  return currentDraftId;
}

async function persistDraftSilentlyIfNeeded() {
  if (!currentDraftId || currentEditId) return;
  try {
    await persistCurrentDraft({ silent: true, source: 'manual' });
  } catch (_) { }
}

function getRejectionHistoryEntries(submission) {
  const rawHistory = Array.isArray(submission?.rejectionHistory) ? submission.rejectionHistory : [];
  const normalizedHistory = rawHistory
    .map((entry) => {
      if (typeof entry === 'string') {
        const reason = entry.trim();
        return reason ? { reason, rejectedAt: null } : null;
      }
      const reason = String(entry?.reason || '').trim();
      if (!reason) return null;
      return {
        reason,
        rejectedAt: entry?.rejectedAt || null,
        rejectedBy: entry?.rejectedBy || entry?.performedBy || entry?.actorEmail || null
      };
    })
    .filter(Boolean);

  if (normalizedHistory.length > 0) return normalizedHistory;

  const auditHistory = Array.isArray(submission?.auditCommissionRejections)
    ? submission.auditCommissionRejections
      .map((entry) => {
        const reason = String(entry?.reason || '').trim();
        if (!reason) return null;
        return {
          reason,
          rejectedAt: entry?.rejectedAt || null,
          rejectedBy: entry?.rejectedBy || entry?.performedBy || null
        };
      })
      .filter(Boolean)
    : [];

  if (auditHistory.length > 0) return auditHistory;

  const auditReason = String(submission?.auditCommissionRejectionReason || '').trim();
  if (auditReason) {
    return [{
      reason: auditReason,
      rejectedAt: submission?.auditCommissionRejectedAt || null,
      rejectedBy: submission?.auditCommissionRejectedBy || null
    }];
  }

  const fallbackReason = String(
    submission?.latestRejectionReason ||
    submission?.previousRejectionReason ||
    submission?.comment ||
    ''
  ).trim();

  return fallbackReason ? [{
    reason: fallbackReason,
    rejectedAt: submission?.latestRejectedAt || submission?.previousRejectedAt || submission?.reviewedAt || null,
    rejectedBy: submission?.latestRejectedBy || submission?.reviewedBy || null
  }] : [];
}

function hasRejectionHistory(submission) {
  return getRejectionHistoryEntries(submission).length > 0;
}

function updateRoleSwitchBackLink() {
  if (!switchBackRoleLink || !switchBackRoleText) return;
  const role = String(currentUserProfile?.role || '').trim().toLowerCase();
  const roleTargets = {
    reviewer: { href: 'reviewer-dashboard.html', label: 'Switch to Reviewer' },
    rsa: { href: 'rsa-dashboard.html', label: 'Switch to RSA' },
    reports_monitoring: { href: 'reports-monitoring-dashboard.html', label: 'Switch to Audit' },
    audit: { href: 'reports-monitoring-dashboard.html', label: 'Switch to Audit' }
  };
  const target = roleTargets[role];
  if (!target) {
    switchBackRoleLink.style.display = 'none';
    switchBackRoleLink.removeAttribute('href');
    return;
  }
  switchBackRoleLink.href = target.href;
  switchBackRoleText.textContent = target.label;
  switchBackRoleLink.style.display = 'flex';
}

function showCompressionPromptModal(file) {
  return new Promise((resolve) => {
    const existingModal = document.getElementById('compressionPromptModal');
    if (existingModal) existingModal.remove();

    const formatSize = (bytes) => (bytes / (1024 * 1024)).toFixed(2) + ' MB';
    const isPdf = (file?.type === 'application/pdf') || String(file?.name || '').toLowerCase().endsWith('.pdf');
    const fileKind = isPdf ? 'PDF' : 'image';

    const modal = document.createElement('div');
    modal.className = 'modal file-size-warning';
    modal.id = 'compressionPromptModal';
    modal.style.cssText = `
      display: flex !important;
      position: fixed !important;
      top: 0 !important;
      left: 0 !important;
      width: 100% !important;
      height: 100% !important;
      background: rgba(0, 0, 0, 0.5) !important;
      z-index: 999999 !important;
      align-items: center !important;
      justify-content: center !important;
      opacity: 1 !important;
    `;

    const modalContent = document.createElement('div');
    modalContent.style.cssText = `
      background: white !important;
      border-radius: 12px !important;
      max-width: 520px !important;
      width: 92% !important;
      max-height: 90vh !important;
      overflow-y: auto !important;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25) !important;
      animation: modalSlideIn 0.3s ease !important;
    `;

    modalContent.innerHTML = `
      <div style="padding: 20px; border-bottom: 1px solid #e2e8f0; background: #eff6ff; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
        <h2 style="margin: 0; color: #1d4ed8; font-size: 20px; display: flex; align-items: center; gap: 10px;">
          <i class="fas fa-file-zipper"></i> Compress File
        </h2>
        <button id="compressionPromptCloseBtn" style="background: transparent; border: none; font-size: 28px; color: #1d4ed8; cursor: pointer; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; border-radius: 50%;">&times;</button>
      </div>
      <div style="padding: 24px;">
        <p style="font-size: 16px; color: #1e293b; margin-bottom: 16px; text-align: center;">
          This ${fileKind} is larger than the ${getPdfUploadLimitLabel()} upload limit.
        </p>
        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 16px; margin-bottom: 16px;">
          <div style="font-weight: 700; color: #0f172a; margin-bottom: 6px;">${file.name}</div>
          <div style="color: #475569;">Current size: <strong>${formatSize(file.size)}</strong></div>
          <div style="color: #475569;">Target size: <strong>${getPdfUploadLimitLabel()} or less</strong></div>
        </div>
        <div style="background: #fefce8; border-radius: 8px; padding: 14px; color: #854d0e; font-size: 13px; line-height: 1.5;">
          Compression keeps the upload inside the system limit. If the file cannot be reduced enough without breaking readability, we will let you know.
        </div>
      </div>
      <div style="padding: 20px; border-top: 1px solid #e2e8f0; display: flex; gap: 10px; justify-content: flex-end; background: #f8fafc; border-radius: 0 0 12px 12px;">
        <button id="compressionPromptCancelBtn" style="background: white; border: 1px solid #cbd5e1; color: #475569; padding: 10px 18px; border-radius: 6px; font-size: 14px; font-weight: 500; cursor: pointer;">
          Cancel
        </button>
        <button id="compressionPromptConfirmBtn" style="background: #003366; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; display:flex; align-items:center; gap:8px;">
          <i class="fas fa-compress-alt"></i> Compress and Continue
        </button>
      </div>
    `;

    modal.appendChild(modalContent);
    (document.body || document.documentElement)?.appendChild(modal);

    const close = (result) => {
      modal.remove();
      resolve(result);
    };

    document.getElementById('compressionPromptCloseBtn')?.addEventListener('click', () => close(false));
    document.getElementById('compressionPromptCancelBtn')?.addEventListener('click', () => close(false));
    document.getElementById('compressionPromptConfirmBtn')?.addEventListener('click', () => close(true));
    modal.addEventListener('click', (e) => {
      if (e.target === modal) close(false);
    });
  });
}

function getDefaultDashboardForRole(role) {
  const normalized = String(role || '').trim().toLowerCase();
  if (normalized === 'reviewer') return 'reviewer-dashboard.html';
  if (normalized === 'rsa') return 'rsa-dashboard.html';
  if (normalized === 'payment') return 'payment-dashboard.html';
  if (normalized === 'admin') return 'admin-dashboard.html';
  if (normalized === 'reports_monitoring' || normalized === 'audit') return 'reports-monitoring-dashboard.html';
  if (normalized === 'super_admin') return 'super-admin-dashboard.html';
  return '';
}

function isExplicitUploaderSwitch() {
  const params = new URLSearchParams(window.location.search);
  return params.get('view') === 'uploader' || params.get('switch') === 'uploader';
}

function syncOriginatingTpField() {
  const originatingTpEl = document.getElementById('originatingTP');
  if (!originatingTpEl) return;
  originatingTpEl.value = String(currentUserProfile?.location || '').trim();
  originatingTpEl.readOnly = true;
  originatingTpEl.setAttribute('readonly', 'readonly');
  originatingTpEl.setAttribute('aria-readonly', 'true');
  originatingTpEl.style.backgroundColor = '#f8fafc';
  originatingTpEl.style.cursor = 'not-allowed';
}

function isCurrentUserRsa() {
  return String(currentUserProfile?.role || '').trim().toLowerCase() === 'rsa';
}

function getCurrentUserRoleLevel() {
  const rawLevel = currentUserProfile?.roleLevel ?? currentUserProfile?.accessLevel ?? 1;
  const normalized = String(rawLevel || '').trim().toLowerCase();
  if (Number(rawLevel) === 2 || normalized === '2' || normalized === 'level 2' || normalized === 'level2') {
    return 2;
  }
  return 1;
}

function isCurrentUserUploaderLevel2() {
  const role = String(currentUserProfile?.role || '').trim().toLowerCase();
  return role === 'uploader' && getCurrentUserRoleLevel() === 2;
}

function canCurrentUserSubmitWithoutDocuments() {
  if (isCurrentUserUploaderLevel2()) return true;
  const role = String(currentUserProfile?.role || '').trim().toLowerCase();
  const level = getCurrentUserRoleLevel();
  if (!role) return false;
  const levelAwareKey = (role === 'uploader' || role === 'rsa') ? `${role}_level_${level}` : role;
  if (Object.prototype.hasOwnProperty.call(currentDocumentRequirementRoles, levelAwareKey)) {
    return currentDocumentRequirementRoles[levelAwareKey] === false;
  }
  return currentDocumentRequirementRoles[role] === false;
}

function syncUploadRequirementUi() {
  const pdfLimit = getPdfUploadLimitLabel();
  if (requiredDocumentsHeading) {
    requiredDocumentsHeading.textContent = canCurrentUserSubmitWithoutDocuments() ? 'Documents' : 'Required Documents';
  }
  if (batchUploadHint) {
    batchUploadHint.textContent = canCurrentUserSubmitWithoutDocuments()
      ? `Documents are optional for your role. Save details first to enable batch upload (Max ${pdfLimit} per file)`
      : `Save details first to enable batch upload (Max ${pdfLimit} per file)`;
  }
  if (uploadProgressText) {
    uploadProgressText.textContent = canCurrentUserSubmitWithoutDocuments() ? ' optional for this role' : '';
  }
  if (submitCustomerBtn) {
    submitCustomerBtn.innerHTML = canCurrentUserSubmitWithoutDocuments()
      ? '<i class="fas fa-check-circle"></i> Submit Application'
      : '<i class="fas fa-check-circle"></i> Submit All Documents';
  }
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', () => {
  auth.onAuthStateChanged(async (user) => {
    if (user) {
      currentUser = user;
      try {
        const data = await getCurrentUserProfileShared(db, user);
        if (data) {
          currentUserProfile = data;
          userName.textContent = data.fullName || user.displayName || user.email;
        } else {
          currentUserProfile = { email: user.email, fullName: user.displayName || user.email, role: 'uploader', status: 'active' };
          userName.textContent = user.displayName || user.email;
        }
      } catch (e) {
        currentUserProfile = { email: user.email, fullName: user.displayName || user.email, role: 'uploader', status: 'active' };
        userName.textContent = user.displayName || user.email;
      }
      const defaultDashboard = getDefaultDashboardForRole(currentUserProfile?.role);
      if (defaultDashboard && !isExplicitUploaderSwitch()) {
        window.location.href = defaultDashboard;
        return;
      }
      userAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3E👤%3C/text%3E%3C/svg%3E';
      renderProfileTab();
      updateRoleSwitchBackLink();
      await applyUploadSystemSettings();
      await applyWorkflowSystemSettings({ force: true });
      await populateAgentBankOptions({ force: true });
      await populateCustomerAccountBankOptions();
      syncUploadRequirementUi();
      await loadRegisteredAgents();
      await loadApprovedAgents();
      await loadSubmissions();
      const initialTab = getInitialUploaderTab();
      const initialApplicationTab = getInitialUploaderApplicationTab();
      window.switchTab(initialTab);
      if (initialTab === 'applications') {
        switchUploaderApplicationTab(initialApplicationTab);
      }
    } else {
      window.location.href = 'index.html';
    }
  });
  setupEventListeners();
});

// ==================== PROPERTY RULES ====================
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
let PROPERTY_RULES = [...DEFAULT_PROPERTY_RULES];
let HOUSE_NUMBER_RULES = { ...DEFAULT_HOUSE_NUMBER_RULES };
let penNoValidationSequence = 0;
let accountLookupSequence = 0;
let accountLookupDebounce = null;
let accountLookupBanks = [];
let agentAccountBankOptions = [];
let verifiedAccountLookup = {
  accountNumber: '',
  bankCode: '',
  accountName: ''
};

const HOUSE_COUNTER_COLLECTION = 'houseNumberCounters';
const HOUSE_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function determinePropertyByRsa(rsaAmount) {
  const n = Number(rsaAmount) || 0;
  for (const r of PROPERTY_RULES) {
    if (n >= r.min && n <= r.max) return r;
  }
  return null;
}

function getHouseNumberRule(propertyType) {
  const normalizedType = String(propertyType || '').trim();
  if (!normalizedType) return null;
  if (HOUSE_NUMBER_RULES[normalizedType]) return HOUSE_NUMBER_RULES[normalizedType];
  const matchedKey = Object.keys(HOUSE_NUMBER_RULES).find((key) => key.toLowerCase() === normalizedType.toLowerCase());
  return matchedKey ? HOUSE_NUMBER_RULES[matchedKey] : null;
}

function houseCounterDocId(propertyType) {
  const key = String(propertyType || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return key ? `house_${key}` : '';
}

function lettersToIndex(value) {
  const text = String(value || '').trim().toUpperCase();
  if (!text) return 0;
  let result = 0;
  for (const char of text) {
    const pos = HOUSE_ALPHABET.indexOf(char);
    if (pos < 0) continue;
    result = (result * 26) + pos + 1;
  }
  return Math.max(0, result - 1);
}

function indexToLetters(index) {
  let value = Number(index);
  if (!Number.isFinite(value) || value < 0) value = 0;
  let text = '';
  let current = Math.floor(value);
  do {
    text = HOUSE_ALPHABET[current % 26] + text;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);
  return text;
}

function formatGeneratedHouseNumber(rule, index) {
  const safeIndex = Math.max(0, Number(index) || 0);
  if (!rule) return '';

  if (rule.mode === 'alpha_suffix') {
    const startLetterIndex = lettersToIndex(rule.startLetter);
    const firstSpan = HOUSE_ALPHABET.length - startLetterIndex;
    if (safeIndex < firstSpan) {
      return `${rule.prefix}${rule.startNumber}${indexToLetters(startLetterIndex + safeIndex)}`;
    }
    const remainder = safeIndex - firstSpan;
    const numberOffset = Math.floor(remainder / HOUSE_ALPHABET.length) + 1;
    const suffixIndex = remainder % HOUSE_ALPHABET.length;
    return `${rule.prefix}${rule.startNumber + numberOffset}${indexToLetters(suffixIndex)}`;
  }

  if (rule.mode === 'block_100') {
    const firstSpan = 101 - Number(rule.startNumber);
    if (safeIndex < firstSpan) {
      return `${rule.startPrefix}${Number(rule.startNumber) + safeIndex}`;
    }
    const remainder = safeIndex - firstSpan;
    const prefixOffset = Math.floor(remainder / 100) + 1;
    const numberValue = (remainder % 100) + 1;
    return `${indexToLetters(lettersToIndex(rule.startPrefix) + prefixOffset)}${numberValue}`;
  }

  if (rule.mode === 'house_infinite') {
    return `House ${Number(rule.startNumber) + safeIndex}`;
  }

  if (rule.mode === 'house_block_100') {
    const firstSpan = 101 - Number(rule.startNumber);
    if (safeIndex < firstSpan) {
      return `House ${rule.startPrefix}${Number(rule.startNumber) + safeIndex}`;
    }
    const remainder = safeIndex - firstSpan;
    const prefixOffset = Math.floor(remainder / 100) + 1;
    const numberValue = (remainder % 100) + 1;
    return `House ${indexToLetters(lettersToIndex(rule.startPrefix) + prefixOffset)}${numberValue}`;
  }

  return '';
}

async function getNextHouseNumberPreview(propertyType) {
  const rule = getHouseNumberRule(propertyType);
  const counterId = houseCounterDocId(propertyType);
  if (!rule || !counterId) return '';
  try {
    const counterRef = doc(db, HOUSE_COUNTER_COLLECTION, counterId);
    const counterSnap = await getDoc(counterRef);
    const lastIndex = counterSnap.exists() ? Number(counterSnap.data()?.lastIndex) : -1;
    return formatGeneratedHouseNumber(rule, Number.isFinite(lastIndex) ? lastIndex + 1 : 0);
  } catch (error) {
    console.warn('Failed to preview house number', error);
    return '';
  }
}

async function reserveHouseNumber(propertyType) {
  const normalizedPropertyType = String(propertyType || '').trim();
  const rule = getHouseNumberRule(normalizedPropertyType);
  const counterId = houseCounterDocId(normalizedPropertyType);
  if (!rule || !counterId) return '';

  const counterRef = doc(db, HOUSE_COUNTER_COLLECTION, counterId);
  try {
    return await runTransaction(db, async (transaction) => {
      const counterSnap = await transaction.get(counterRef);
      const lastIndex = counterSnap.exists() ? Number(counterSnap.data()?.lastIndex) : -1;
      const nextIndex = Number.isFinite(lastIndex) ? lastIndex + 1 : 0;
      const houseNumber = formatGeneratedHouseNumber(rule, nextIndex);
      transaction.set(counterRef, {
        propertyType: normalizedPropertyType,
        lastIndex: nextIndex,
        lastHouseNumber: houseNumber,
        updatedAt: serverTimestamp()
      }, { merge: true });
      return houseNumber;
    });
  } catch (error) {
    console.warn('Failed to reserve house number', error);
    return '';
  }
}

async function refreshHouseNumberPreview(options = {}) {
  const houseNumberEl = document.getElementById('houseNumber');
  const propertyType = document.getElementById('propertyType')?.value?.trim() || '';
  if (!houseNumberEl) return '';
  if (!propertyType) {
    houseNumberEl.value = '';
    return '';
  }
  if (currentEditId && houseNumberEl.value.trim() && !options.force) {
    return houseNumberEl.value.trim();
  }
  const preview = await getNextHouseNumberPreview(propertyType);
  houseNumberEl.value = preview;
  return preview;
}

// ==================== UPDATED: SET PROPERTY BY RSA WITH AUTO-FILL ====================
function setPropertyByRsaAndUpdate(rsaAmount) {
  const rule = determinePropertyByRsa(rsaAmount);
  const propEl = document.getElementById('propertyType');
  const propValEl = document.getElementById('propertyValue');
  const propValFmt = document.getElementById('propertyValueFormatted');
  const feeEl = document.getElementById('facilityFee');
  const feeFmt = document.getElementById('facilityFeeFormatted');
  const loanEl = document.getElementById('loanAmount');
  const loanFmt = document.getElementById('loanAmountFormatted');
  const tenorEl = document.getElementById('tenor');

  if (rule) {
    // Set property type
    if (propEl) propEl.value = rule.name;

    // Set property value
    if (propValEl) propValEl.value = rule.value;
    if (propValFmt) propValFmt.textContent = formatCurrency(rule.value);

    // Set facility fee
    if (feeEl) feeEl.value = rule.fee;
    if (feeFmt) feeFmt.textContent = formatCurrency(rule.fee);

    // Calculate 25% of RSA rounded down to the nearest thousand.
    const rsaBalance = parseFloat(rsaAmount) || 0;
    const rsa25Rounded = calculateRoundedRsa25(rsaBalance);

    // Property value must equal loan amount + rounded 25% contribution.
    const loanAmount = roundUpToNearestThousand(rule.value - rsa25Rounded);

    // Set loan amount
    if (loanEl) loanEl.value = loanAmount;
    if (loanFmt) loanFmt.textContent = formatCurrency(loanAmount);

    // Update tenor if DOB is available
    const dob = document.getElementById('customerDob')?.value;
    if (dob && window.calculateTenorFromDob) {
      const tenor = window.calculateTenorFromDob(dob);
      if (tenorEl) tenorEl.value = tenor;
    }

    // Also update the 25% display
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) {
      rsa25FormattedEl.textContent = formatCurrency(rsa25Rounded);
    }

    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) {
      rsa25PercentEl.value = rsa25Rounded;
    }

    void refreshHouseNumberPreview();

  } else {
    // Clear all fields if no rule matches
    if (propEl) propEl.value = '';
    if (propValEl) propValEl.value = '';
    if (propValFmt) propValFmt.textContent = '';
    if (feeEl) feeEl.value = '';
    if (feeFmt) feeFmt.textContent = '';
    if (loanEl) loanEl.value = '';
    if (loanFmt) loanFmt.textContent = '';

    // Also clear 25% display
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) {
      rsa25FormattedEl.textContent = '';
    }

    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) {
      rsa25PercentEl.value = '';
    }
    const houseNumberEl = document.getElementById('houseNumber');
    if (houseNumberEl) {
      houseNumberEl.value = '';
    }
  }
}

window.determinePropertyByRsa = determinePropertyByRsa;
window.setPropertyByRsaAndUpdate = setPropertyByRsaAndUpdate;

window.calculateTenorFromDob = function(dobString) {
  if (!dobString) return '';
  const birthDate = new Date(dobString);
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  const tenor = 60 - age;
  return isNaN(tenor) ? '' : Math.max(0, tenor);
};

// ==================== CALCULATE AND DISPLAY CUSTOMER INFO ====================
function calculateAndDisplayCustomerInfo() {
  const dob = document.getElementById('customerDob')?.value;
  const nin = document.getElementById('customerNIN')?.value?.trim() || '';
  const rsaBalance = parseFloat(document.getElementById('rsaBalance')?.value || 0);

  if (!dob) return;

  // Calculate age
  const birthDate = new Date(dob);
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }

  // Check if adult (age >= 18)
  const isAdult = age >= 18;

  // Check NIN validity (simple check - 11 digits)
  const ninValid = /^\d{11}$/.test(nin);

  // Calculate years to retirement (assuming retirement at 60)
  const yearsToRetirement = Math.max(0, 60 - age);

  // Get 25% of RSA rounded down to the nearest thousand.
  const rsa25Rounded = calculateRoundedRsa25(rsaBalance);

  // Get property rule
  const rule = determinePropertyByRsa(rsaBalance);
  const propertyValue = rule ? rule.value : 0;
  const loanAmount = roundUpToNearestThousand(propertyValue - rsa25Rounded);

  // Find or create results container
  let resultsContainer = document.getElementById('customerInfoResults');
  if (!resultsContainer) {
    resultsContainer = document.createElement('div');
    resultsContainer.id = 'customerInfoResults';
    resultsContainer.style.cssText = 'margin-top: 20px; background: #f0f9ff; padding: 15px; border-radius: 8px; border-left: 4px solid #003366;';

    // Insert after the form actions
    const formActions = document.querySelector('.form-actions') || document.getElementById('saveDetailsBtn')?.parentNode;
    if (formActions && formActions.parentNode) {
      formActions.parentNode.insertBefore(resultsContainer, formActions.nextSibling);
    } else {
      // Fallback - append to customer details section
      const detailsSection = document.querySelector('.customer-details-section');
      if (detailsSection) detailsSection.appendChild(resultsContainer);
    }
  }

  // Update results
  resultsContainer.innerHTML = `
    <h4 style="margin: 0 0 12px 0; color: #003366; font-size: 16px; display: flex; align-items: center; gap: 8px;">
      <i class="fas fa-chart-line"></i> Customer Analysis
    </h4>
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px;">
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Age</div>
        <div style="font-size: 18px; font-weight: 700; color: #003366;">${age} years</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Adult</div>
        <div style="font-size: 18px; font-weight: 700; color: ${isAdult ? '#10b981' : '#ef4444'};">${isAdult ? 'Yes ✓' : 'No ✗'}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">NIN Valid</div>
        <div style="font-size: 18px; font-weight: 700; color: ${ninValid ? '#10b981' : '#f59e0b'};">${ninValid ? 'Valid ✓' : 'Check Required'}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Years to Retirement</div>
        <div style="font-size: 18px; font-weight: 700; color: #003366;">${yearsToRetirement}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">25% of RSA</div>
        <div style="font-size: 18px; font-weight: 700; color: #003366;">${formatCurrency(rsa25Rounded)}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Loan Amount</div>
        <div style="font-size: 18px; font-weight: 700; color: #10b981;">${formatCurrency(loanAmount)}</div>
      </div>
    </div>
  `;

  resultsContainer.style.display = 'block';
}

// ==================== SAVE CUSTOMER DETAILS ====================
async function saveCustomerDetails() {
  // Define requiredFields INSIDE the function
  const allowOptionalCustomerFields = isCurrentUserUploaderLevel2();
  const requiredFields = [
    'customerName', 'customerDob', 'customerEmail', 'customerPhone',
    'customerNIN', 'customerAddress', 'accountNo', 'employer',
    'originatingTP', 'mortgageLoanApplicationFormDate', 'pfa', 'penNo', 'rsaStatementDate', 'rsaBalance'
  ];

  if (String(document.getElementById('accountNo')?.value || '').trim()) {
    if (!(await resolveCustomerAccountName({ silent: false }))) {
      return false;
    }
  }

  const missingFields = [];
  const invalidFields = [];

  requiredFields.forEach(id => {
    const el = document.getElementById(id);
    const value = String(el?.value || '').trim();
    const isRequiredForUser = !allowOptionalCustomerFields || id === 'customerName';
    if (!el || !value) {
      if (!isRequiredForUser) return;
      missingFields.push(id);
      return;
    }
    if (id === 'accountNo' && !/^\d{10}$/.test(value)) invalidFields.push('Account Number (must be exactly 10 digits)');
    if (id === 'customerPhone' && !/^\d{11}$/.test(value)) invalidFields.push('Phone (must be exactly 11 digits)');
    if (id === 'customerNIN' && !/^\d{11}$/.test(value)) invalidFields.push('NIN (must be exactly 11 digits)');
  });

  if (missingFields.length > 0) {
    showNotification('Please fill all required fields before saving', 'error');
    return false;
  }

  if (invalidFields.length > 0) {
    showNotification('Invalid fields: ' + invalidFields.join(', '), 'error');
    return false;
  }

  const penNo = document.getElementById('penNo')?.value?.trim() || '';
  const duplicateExcludeId = currentEditId || currentDraftId || '';
  if (penNo && !(await validatePenNumberAvailable(penNo, duplicateExcludeId, { inline: true, notify: false }))) {
    return false;
  }

  const rsaBalance = parseFloat(document.getElementById('rsaBalance')?.value || 0);

  // Calculate 25% rounded down to the nearest thousand.
  const rsa25Rounded = calculateRoundedRsa25(rsaBalance);

  // Update RSA 25% display with rounded value.
  const rsa25FormattedEl = document.getElementById('rsa25Formatted');
  if (rsa25FormattedEl) {
    rsa25FormattedEl.textContent = formatCurrency(rsa25Rounded);
  }

  // Store rounded value in hidden input.
  const rsa25PercentEl = document.getElementById('rsa25Percent');
  if (rsa25PercentEl) {
    rsa25PercentEl.value = rsa25Rounded;
  }

  // Call property update function
  if (typeof setPropertyByRsaAndUpdate === 'function') {
    setPropertyByRsaAndUpdate(rsaBalance);
  } else if (window.setPropertyByRsaAndUpdate) {
    window.setPropertyByRsaAndUpdate(rsaBalance);
  }

  const dob = document.getElementById('customerDob')?.value;
  if (dob) {
    const tenor = window.calculateTenorFromDob ? window.calculateTenorFromDob(dob) : '';
    const tenorEl = document.getElementById('tenor');
    if (tenorEl) {
      tenorEl.value = tenor;
    }
  }

  // Calculate and display customer info
  if (typeof calculateAndDisplayCustomerInfo === 'function') {
    calculateAndDisplayCustomerInfo();
  }

  await refreshHouseNumberPreview({ force: !currentEditId });
  const reservedDetails = await ensureStoredHouseNumber(collectDraftableCustomerDetails());
  const propertyTypeEl = document.getElementById('propertyType');
  if (propertyTypeEl && reservedDetails.propertyType) {
    propertyTypeEl.value = reservedDetails.propertyType;
  }

  customerDetailsSaved = true;
  showNotification('✅ Customer details saved successfully!', 'success');

  if (batchUploadBtn) {
    batchUploadBtn.disabled = false;
  }

  // Enable document grid
  if (documentGrid) {
    documentGrid.style.pointerEvents = 'auto';
    documentGrid.style.opacity = '1';
  }
  if (optionalDocumentGrid) {
    optionalDocumentGrid.style.pointerEvents = 'auto';
    optionalDocumentGrid.style.opacity = '1';
  }

  if (saveDetailsBtn) {
    saveDetailsBtn.innerHTML = '<i class="fas fa-check"></i> Details Saved';
  }

  updateSubmitButton();
  syncUploadRequirementUi();

  const firstDocumentSection = document.querySelector('.document-grid-section');
  if (firstDocumentSection) {
    firstDocumentSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  await persistDraftSilentlyIfNeeded();

  return true;
}

function resetCustomerDetails() {
  const fields = [
    'customerName', 'customerDob', 'customerEmail', 'customerPhone',
    'customerNIN', 'customerAddress', 'accountNo', 'accountBank', 'employer',
    'originatingTP', 'mortgageLoanApplicationFormDate', 'pfa', 'penNo', 'rsaStatementDate', 'rsaBalance',
    'propertyType', 'houseNumber', 'tenor'
  ];

  // Reset all text fields
  fields.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  if (accountBankSelect) accountBankSelect.value = DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME;
  clearPenNoError();
  clearVerifiedAccountLookup();
  setAccountLookupStatus('', 'info');

  syncOriginatingTpField();

  // Reset RSA 25% fields with null checks
  const rsa25PercentEl = document.getElementById('rsa25Percent');
  if (rsa25PercentEl) rsa25PercentEl.value = '';

  const rsa25FormattedEl = document.getElementById('rsa25Formatted');
  if (rsa25FormattedEl) rsa25FormattedEl.textContent = '';

  // Reset property value fields
  const propertyValueEl = document.getElementById('propertyValue');
  if (propertyValueEl) propertyValueEl.value = '';

  const propertyValueFormattedEl = document.getElementById('propertyValueFormatted');
  if (propertyValueFormattedEl) propertyValueFormattedEl.textContent = '';

  // Reset facility fee fields
  const facilityFeeEl = document.getElementById('facilityFee');
  if (facilityFeeEl) facilityFeeEl.value = '';

  const facilityFeeFormattedEl = document.getElementById('facilityFeeFormatted');
  if (facilityFeeFormattedEl) facilityFeeFormattedEl.textContent = '';

  // Reset loan amount fields
  const loanAmountEl = document.getElementById('loanAmount');
  if (loanAmountEl) loanAmountEl.value = '';

  const loanAmountFormattedEl = document.getElementById('loanAmountFormatted');
  if (loanAmountFormattedEl) loanAmountFormattedEl.textContent = '';

  // Hide results container
  const resultsContainer = document.getElementById('customerInfoResults');
  if (resultsContainer) {
    resultsContainer.style.display = 'none';
  }

  // Reset state
  customerDetailsSaved = false;
  currentHouseNumberReserved = false;
  if (saveDetailsBtn) {
    saveDetailsBtn.innerHTML = '<i class="fas fa-arrow-right"></i> Next: Upload Documents';
  }

  // Disable batch upload button
  if (batchUploadBtn) {
    batchUploadBtn.disabled = true;
  }

  // Disable document grid
  if (documentGrid) {
    documentGrid.style.pointerEvents = 'none';
    documentGrid.style.opacity = '0.5';
  }
  if (optionalDocumentGrid) {
    optionalDocumentGrid.style.pointerEvents = 'none';
    optionalDocumentGrid.style.opacity = '0.5';
  }

  showNotification('Form reset', 'info');
}

function validateDetailsBeforeBatch() {
  if (!customerDetailsSaved) {
    showNotification('Please save customer details before batch uploading', 'error');
    return false;
  }

  const customerNameInput = document.getElementById('customerName');
  const customerName = customerNameInput ? customerNameInput.value.trim() : '';

  if (!customerName) {
    showNotification('Customer name is required', 'error');
    return false;
  }

  return true;
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
  if (newUploadBtn) newUploadBtn.addEventListener('click', openUploadModal);
  if (bulkTemplateBtn) bulkTemplateBtn.addEventListener('click', downloadBulkImportTemplate);
  if (bulkImportBtn && bulkImportInput) {
    bulkImportBtn.addEventListener('click', () => bulkImportInput.click());
    bulkImportInput.addEventListener('change', handleBulkImportSelection);
  }
  if (helpFab && sopHelpModal) {
    helpFab.addEventListener('click', () => sopHelpModal.classList.add('active'));
  }
  if (closeSopHelpModalBtn && sopHelpModal) {
    closeSopHelpModalBtn.addEventListener('click', () => sopHelpModal.classList.remove('active'));
  }
  if (closeSopHelpModalFooterBtn && sopHelpModal) {
    closeSopHelpModalFooterBtn.addEventListener('click', () => sopHelpModal.classList.remove('active'));
  }
  document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
    item.addEventListener('click', (e) => { e.preventDefault(); switchTab(item.dataset.tab); });
  });
  document.querySelectorAll('[data-uploader-application-tab]').forEach((button) => {
    button.addEventListener('click', () => switchUploaderApplicationTab(button.dataset.uploaderApplicationTab || 'pending'));
  });
  document.querySelectorAll('[data-active-commission-scope]').forEach((button) => {
    button.addEventListener('click', () => {
      currentUploaderPaidScope = button.dataset.activeCommissionScope === 'others' ? 'others' : 'mine';
      renderUploaderApplicationsTable();
    });
  });
  if (applicationsSearch) {
    applicationsSearch.addEventListener('input', renderUploaderApplicationsTable);
  }
  if (rejectedTableBody) {
    rejectedTableBody.addEventListener('click', (e) => {
      const chatTrigger = e.target.closest('.app-chat-trigger');
      if (chatTrigger) return;
      const waLink = e.target.closest('a[href*="wa.me/"]');
      if (!waLink) return;
      const row = waLink.closest('tr');
      const submissionId = row?.getAttribute('data-submission-id');
      if (!submissionId) return;
      e.preventDefault();
      e.stopPropagation();
      if (typeof window.openApplicationChat === 'function') {
        window.openApplicationChat(submissionId);
      }
    });
  }
  if (closeModalBtn) closeModalBtn.addEventListener('click', () => closeModal(uploadModal));
  if (closeEditModalBtn) closeEditModalBtn.addEventListener('click', () => closeModal(editModal));
  if (closeSingleModalBtn) closeSingleModalBtn.addEventListener('click', () => closeModal(singleUploadModal));
  if (closeErrorDetailModalBtn) closeErrorDetailModalBtn.addEventListener('click', () => closeModal(errorDetailModal));
  if (dismissErrorDetailModalBtn) dismissErrorDetailModalBtn.addEventListener('click', () => closeModal(errorDetailModal));
  if (closeUploaderRejectionReasonModal) closeUploaderRejectionReasonModal.addEventListener('click', () => closeModal(uploaderRejectionReasonModal));
  if (closeUploaderRejectionReasonBtn) closeUploaderRejectionReasonBtn.addEventListener('click', () => closeModal(uploaderRejectionReasonModal));
  if (cancelSingleBtn) cancelSingleBtn.addEventListener('click', () => closeModal(singleUploadModal));
  if (submitCustomerBtn) submitCustomerBtn.addEventListener('click', submitCustomer);
  if (saveDraftBtn) saveDraftBtn.addEventListener('click', async () => {
    if (saveDraftBtn.disabled) return;
    saveDraftBtn.disabled = true;
    const originalHtml = saveDraftBtn.innerHTML;
    saveDraftBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
    try {
      await persistCurrentDraft({ silent: false, source: 'manual' });
    } catch (error) {
      showNotification('Draft save failed: ' + error.message, 'error');
    } finally {
      saveDraftBtn.disabled = false;
      saveDraftBtn.innerHTML = originalHtml;
    }
  });
  if (submitEditBtn) submitEditBtn.addEventListener('click', submitEdit);
  if (singleUploadArea) {
    setupDragDrop(singleUploadArea, singleFileInput, async (files) => {
      if (files.length > 0) await handleSingleFileSelection(files[0]);
    });
  }
  if (singleFileInput) {
    singleFileInput.addEventListener('change', async (e) => {
      if (e.target.files.length > 0) await handleSingleFileSelection(e.target.files[0]);
    });
  }
  if (batchUploadBtn && batchFileInput) {
    batchUploadBtn.disabled = true;
    batchUploadBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (validateDetailsBeforeBatch()) batchFileInput.click();
    });
    batchFileInput.addEventListener('change', (e) => {
      const files = Array.from(e.target.files);
      if (files.length > 0) {
        const oversizedFiles = files.filter((file) => file.size > MAX_PDF_SIZE_BYTES);
        const validFiles = files.filter((file) => file.size <= MAX_PDF_SIZE_BYTES);
        if (oversizedFiles.length > 0) {
          showBatchSizeWarningModal(oversizedFiles, validFiles, files);
        } else {
          prepareBatchForMapping(files);
        }
        batchFileInput.value = '';
      }
    });
  }
  // FIX: Don't exit early if confirmSingleUpload is missing - still show modal
  if (confirmSingleUpload) {
    confirmSingleUpload.addEventListener('click', uploadSingleDocument);
  }
  document.querySelectorAll('.browse-text').forEach(text => {
    text.addEventListener('click', (e) => {
      const input = e.target.closest('.upload-area').querySelector('input[type="file"]');
      if (input) input.click();
    });
  });
  if (customerNameInput) customerNameInput.addEventListener('input', updateSubmitButton);
  if (accountNoInput) {
    accountNoInput.addEventListener('input', () => {
      const digits = String(accountNoInput.value || '').replace(/\D/g, '').slice(0, 10);
      accountNoInput.value = digits;
      clearVerifiedAccountLookup();
      setAccountLookupStatus('', 'info');
      scheduleCustomerAccountLookup();
    });
    accountNoInput.addEventListener('blur', () => {
      void resolveCustomerAccountName();
    });
  }
  if (accountBankSelect) {
    accountBankSelect.addEventListener('change', () => {
      clearVerifiedAccountLookup();
      setAccountLookupStatus('', 'info');
      void resolveCustomerAccountName();
    });
  }
  if (penNoInput) {
    penNoInput.addEventListener('input', () => {
      penNoValidationSequence += 1;
      clearPenNoError();
    });
    penNoInput.addEventListener('blur', () => {
      void validatePenNoFieldInline();
    });
  }
  if (saveDetailsBtn) saveDetailsBtn.addEventListener('click', saveCustomerDetails);
  if (resetDetailsBtn) resetDetailsBtn.addEventListener('click', resetCustomerDetails);
  if (agentRegistrationForm) {
    agentRegistrationForm.addEventListener('submit', handleAgentRegistration);
  }
  if (openAgentRegistrationBtn) {
    openAgentRegistrationBtn.addEventListener('click', openAgentRegistrationModal);
  }
  if (closeAgentRegistrationModalBtn) {
    closeAgentRegistrationModalBtn.addEventListener('click', closeAgentRegistrationModal);
  }
  if (resetAgentFormBtn) {
    resetAgentFormBtn.addEventListener('click', () => {
      agentRegistrationForm?.reset();
      clearVerifiedAgentAccountLookup();
      setAgentAccountLookupStatus('', 'info');
    });
  }
  if (agentAccountNumberInput) {
    agentAccountNumberInput.addEventListener('input', () => {
      const digits = String(agentAccountNumberInput.value || '').replace(/\D/g, '').slice(0, 10);
      agentAccountNumberInput.value = digits;
      clearVerifiedAgentAccountLookup();
      setAgentAccountLookupStatus('', 'info');
      if (digits.length === 10 && getSelectedAgentAccountBank().code) {
        scheduleAgentAccountLookup();
      }
    });
    agentAccountNumberInput.addEventListener('blur', () => {
      if (getSelectedAgentAccountBank().code) {
        void resolveAgentAccountName();
      }
    });
  }
  if (agentAccountBankSelect) {
    agentAccountBankSelect.addEventListener('input', () => {
      clearVerifiedAgentAccountLookup();
      setAgentAccountLookupStatus('', 'info');
      const bank = getSelectedAgentAccountBank();
      const accountNumber = String(agentAccountNumberInput?.value || '').replace(/\D/g, '');
      if (bank.code && accountNumber.length === 10) {
        const matchedBank = findAgentAccountBankByValue(agentAccountBankSelect.value);
        const suggestedName = String(matchedBank?.accountName || '').trim();
        if (suggestedName) {
          verifiedAgentAccountLookup = { accountNumber, bankCode: bank.code, accountName: suggestedName };
          if (agentAccountNameInput) agentAccountNameInput.value = suggestedName;
          setAgentAccountLookupStatus(`Verified: ${suggestedName}`, 'success');
        } else {
          void resolveAgentAccountName();
        }
      }
    });
    agentAccountBankSelect.addEventListener('change', () => {
      clearVerifiedAgentAccountLookup();
      setAgentAccountLookupStatus('', 'info');
      void resolveAgentAccountName();
    });
  }

  // ===== UPDATED: RSA Balance input listener with auto-fill =====
  document.getElementById('rsaBalance')?.addEventListener('input', (e) => {
    const rawStr = String(e.target.value).replace(/[^0-9.\-]+/g, '');
    const raw = parseFloat(rawStr) || 0;

    // Calculate 25% rounded down to the nearest thousand.
    const rsa25Rounded = calculateRoundedRsa25(raw);

    // Update RSA 25% display immediately
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) {
      rsa25FormattedEl.textContent = formatCurrency(rsa25Rounded);
    }

    // Store rounded value in hidden input.
    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) {
      rsa25PercentEl.value = rsa25Rounded;
    }

    // Call property update function to auto-fill all property fields
    if (typeof setPropertyByRsaAndUpdate === 'function') {
      setPropertyByRsaAndUpdate(raw);
    } else if (window.setPropertyByRsaAndUpdate) {
      window.setPropertyByRsaAndUpdate(raw);
    }

    // Update customer info display
    if (typeof calculateAndDisplayCustomerInfo === 'function') {
      calculateAndDisplayCustomerInfo();
    }
  });

  // ===== UPDATED: DOB change listener to update tenor =====
  document.getElementById('customerDob')?.addEventListener('change', (e) => {
    const dob = e.target.value;
    const t = window.calculateTenorFromDob ? window.calculateTenorFromDob(dob) : '';
    const tenorEl = document.getElementById('tenor');
    if (tenorEl) tenorEl.value = t;

    // Also recalculate property fields if RSA balance exists
    const rsaBalance = document.getElementById('rsaBalance')?.value;
    if (rsaBalance && rsaBalance.trim() !== '') {
      const rsaVal = parseFloat(rsaBalance) || 0;
      if (typeof setPropertyByRsaAndUpdate === 'function') {
        setPropertyByRsaAndUpdate(rsaVal);
      } else if (window.setPropertyByRsaAndUpdate) {
        window.setPropertyByRsaAndUpdate(rsaVal);
      }
    }
  });

  if (document.getElementById('closeViewer')) {
    document.getElementById('closeViewer').addEventListener('click', closeViewerModal);
  }
  window.addEventListener('click', (e) => {
    if (e.target === uploadModal) closeModal(uploadModal);
    if (e.target === editModal) closeModal(editModal);
    if (e.target === singleUploadModal) closeModal(singleUploadModal);
    if (e.target === errorDetailModal) closeModal(errorDetailModal);
    if (e.target === uploaderRejectionReasonModal) closeModal(uploaderRejectionReasonModal);
    if (e.target === viewerModal) closeViewerModal();
    if (e.target === sopHelpModal) sopHelpModal.classList.remove('active');
  });
  if (closeBatchMappingBtn) closeBatchMappingBtn.addEventListener('click', () => batchMappingModal.classList.remove('active'));
  if (cancelBatchMapping) cancelBatchMapping.addEventListener('click', () => batchMappingModal.classList.remove('active'));
  if (confirmBatchMapping) {
    confirmBatchMapping.addEventListener('click', async () => {
      const mapped = __batchFilesBuffer.filter(b => b.selectedType).map(b => ({ file: b.originalFile, mappedType: b.selectedType }));
      if (mapped.length === 0) { showNotification('No files selected for upload', 'error'); return; }
      batchMappingModal.classList.remove('active');
      await handleBatchFiles(mapped);
    });
  }
  try {
    const rsaVal = parseFloat(document.getElementById('rsaBalance')?.value || 0);
    if (rsaVal && window.setPropertyByRsaAndUpdate) window.setPropertyByRsaAndUpdate(rsaVal);
    const dob = document.getElementById('customerDob')?.value;
    if (dob && window.calculateTenorFromDob) {
      const t = window.calculateTenorFromDob(dob);
      const tenorEl = document.getElementById('tenor'); if (tenorEl) tenorEl.value = t;
    }
  } catch (e) { }

}

// ==================== TAB SWITCHING ====================
window.switchTab = (tabId) => {
  tabId = UPLOADER_DASHBOARD_TABS.includes(tabId) ? tabId : 'overview';
  rememberUploaderTab(tabId, tabId === 'applications' ? currentUploaderApplicationTab : '');
  document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
  document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
  document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
  document.getElementById(`${tabId}Tab`)?.classList.add('active');
  const titles = { overview: 'Dashboard', draft: 'Draft Submissions', applications: 'Applications', pending: 'Pending Documents', approved: 'Approved Documents', rejected: 'Rejected Documents', paid: 'Commission', 'register-agent': 'Register Agent', profile: 'My Profile' };
  if (pageTitle) pageTitle.textContent = titles[tabId] || 'My Documents';
  if (tabId === 'applications') {
    renderUploaderApplicationsTable();
  }
  if (tabId === 'paid') {
    renderPaidTable();
  }
  if (tabId === 'register-agent') {
    loadApprovedAgents();
  }
};

function switchCommissionTab(tab) {
  currentCommissionTab = tab === 'cleared' ? 'cleared' : tab === 'active' ? 'active' : 'sent_to_pfa';
  document.getElementById('uploaderAgentCommissionSentTabBtn')?.classList.toggle('active', currentCommissionTab === 'sent_to_pfa');
  document.getElementById('uploaderAgentCommissionActiveTabBtn')?.classList.toggle('active', currentCommissionTab === 'active');
  document.getElementById('uploaderAgentCommissionClearedTabBtn')?.classList.toggle('active', currentCommissionTab === 'cleared');
  renderUploaderAgentCommissionBreakdown(currentCommissionTab);
}

function getDateFromAny(value) {
  if (!value) return null;
  try {
    if (value.toDate && typeof value.toDate === 'function') return value.toDate();
    if (value.seconds) return new Date(value.seconds * 1000);
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  } catch (_) {
    return null;
  }
}

function inDateRange(dateObj, startStr, endStr) {
  if (!dateObj) return false;
  if (!startStr && !endStr) return true;

  if (startStr) {
    const start = new Date(startStr);
    start.setHours(0, 0, 0, 0);
    if (dateObj < start) return false;
  }
  if (endStr) {
    const end = new Date(endStr);
    end.setHours(23, 59, 59, 999);
    if (dateObj > end) return false;
  }
  return true;
}

// ==================== DRAG & DROP ====================
function setupDragDrop(area, input, callback) {
  area.addEventListener('dragover', (e) => { e.preventDefault(); area.classList.add('drag-over'); });
  area.addEventListener('dragleave', () => area.classList.remove('drag-over'));
  area.addEventListener('drop', (e) => { e.preventDefault(); area.classList.remove('drag-over'); callback(e.dataTransfer.files); });
}

// ==================== MODAL FUNCTIONS ====================
function openUploadModal() {
  currentEditId = null;
  currentDraftId = null;
  currentHouseNumberReserved = false;
  currentCustomerUploads = {};
  currentDocType = null;
  currentFile = null;
  setUploadModalHeading('new');
  syncUploadModalControlState('new');

  if (singleFileInput) singleFileInput.value = '';
  if (saveDetailsBtn) {
    saveDetailsBtn.innerHTML = '<i class="fas fa-arrow-right"></i> Next: Upload Documents';
  }
  if (confirmSingleUpload) {
    confirmSingleUpload.disabled = true;
    confirmSingleUpload.innerHTML = 'Upload Document';
  }
  clearSingleFilePreview();

  const fieldsToClear = [
    'customerName', 'customerDob', 'customerEmail', 'customerPhone',
    'customerNIN', 'customerAddress', 'accountNo', 'accountBank', 'employer',
    'originatingTP', 'mortgageLoanApplicationFormDate', 'pfa', 'penNo',
    'rsaStatementDate', 'rsaBalance', 'propertyType', 'propertyValue',
    'facilityFee', 'loanAmount', 'tenor'
  ];

  fieldsToClear.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  if (accountBankSelect) accountBankSelect.value = DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME;
  clearVerifiedAccountLookup();
  setAccountLookupStatus('', 'info');

  const rsa25PercentEl = document.getElementById('rsa25Percent');
  if (rsa25PercentEl) rsa25PercentEl.value = '';

  const rsa25FormattedEl = document.getElementById('rsa25Formatted');
  if (rsa25FormattedEl) rsa25FormattedEl.textContent = '';

  const propertyValueFormattedEl = document.getElementById('propertyValueFormatted');
  if (propertyValueFormattedEl) propertyValueFormattedEl.textContent = '';

  const facilityFeeFormattedEl = document.getElementById('facilityFeeFormatted');
  if (facilityFeeFormattedEl) facilityFeeFormattedEl.textContent = '';

  const loanAmountFormattedEl = document.getElementById('loanAmountFormatted');
  if (loanAmountFormattedEl) loanAmountFormattedEl.textContent = '';

  const resultsContainer = document.getElementById('customerInfoResults');
  if (resultsContainer) resultsContainer.style.display = 'none';

  currentSubmissionAgentFallback = null;
  if (customerAgentSelect) customerAgentSelect.value = '';
  customerDetailsSaved = false;

  syncOriginatingTpField();

  if (batchUploadBtn) batchUploadBtn.disabled = true;
  if (window.renderDocumentGridUpload) {
    window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'upload');
    window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'upload');
  }
  // New submission: lock document grids until details are saved.
  if (documentGrid) { documentGrid.style.pointerEvents = 'none'; documentGrid.style.opacity = '0.5'; }
  if (optionalDocumentGrid) { optionalDocumentGrid.style.pointerEvents = 'none'; optionalDocumentGrid.style.opacity = '0.5'; }
  updateSubmitButton();
  uploadModal.classList.add('active');
}

function closeModal(modal) {
  if (!modal) return;
  modal.classList.remove('active');
  if (modal === singleUploadModal) {
    currentFile = null;
    clearSingleFilePreview();
    if (confirmSingleUpload) confirmSingleUpload.disabled = true;
    if (singleFileInput) singleFileInput.value = '';
  } else if (modal === uploaderRejectionReasonModal) {
    if (uploaderRejectionReasonCustomerName) uploaderRejectionReasonCustomerName.textContent = '-';
    if (uploaderRejectionReasonContact) uploaderRejectionReasonContact.textContent = 'Contact: -';
    if (uploaderRejectionReasonHistory) {
      uploaderRejectionReasonHistory.innerHTML = '';
      uploaderRejectionReasonHistory.style.display = 'none';
    }
  }
}

function showErrorDetailModal(message) {
  if (!errorDetailModal || !errorDetailMessage) return;
  errorDetailMessage.textContent = String(message || 'An unexpected error occurred.');
  errorDetailModal.classList.add('active');
}

function closeViewerModal() {
  if (viewerModal) viewerModal.classList.remove('active');
  if (documentViewer) documentViewer.src = '';
  const viewerHeader = document.querySelector('.viewer-header');
  if (viewerHeader) {
    const existingNav = viewerHeader.querySelector('.viewer-nav');
    if (existingNav) existingNav.remove();
  }
}

// ==================== RENDER DOCUMENT GRID ====================
function renderDocumentGrid(container, documentTypes, uploadedDocs, mode) {
  if (!container) return;
  container.innerHTML = documentTypes.map(doc => {
    const isUploaded = uploadedDocs[doc.id] && uploadedDocs[doc.id].length > 0;
    const statusEmoji = isUploaded ? '✅' : '⬜';
    const statusClass = isUploaded ? 'uploaded' : 'pending';
    return `
      <div class="document-grid-item ${statusClass}" data-doc-type="${doc.id}">
        <div class="doc-status"><span class="status-emoji">${statusEmoji}</span></div>
        <div class="doc-icon"><i class="fas ${doc.icon}"></i></div>
        <div class="doc-name">${doc.name}</div>
        ${isUploaded ? `
          <div class="doc-actions">
            <button class="doc-btn view-btn" onclick="window.viewDocument('${doc.id}')"><i class="fas fa-eye"></i> View</button>
            ${mode === 'edit' ? `
              <button class="doc-btn upload-btn" onclick="window.openSingleUpload('${doc.id}', '${doc.name}')"><i class="fas fa-upload"></i> Re-upload</button>
              <button class="doc-btn remove-btn" onclick="window.removeUploadedDoc('${doc.id}')"><i class="fas fa-trash"></i> Remove</button>
            ` : `<button class="doc-btn remove-btn" onclick="window.removeUploadedDoc('${doc.id}')"><i class="fas fa-trash"></i> Remove</button>`}
          </div>
        ` : `<div class="doc-actions"><button class="doc-btn upload-btn" onclick="window.openSingleUpload('${doc.id}', '${doc.name}')"><i class="fas fa-upload"></i> Upload</button></div>`}
      </div>
    `;
  }).join('');
}
window.renderDocumentGridUpload = renderDocumentGrid;

// ==================== SINGLE UPLOAD ====================
window.openSingleUpload = (docType, docName) => {
  if (!customerDetailsSaved && !currentEditId) { showNotification('Please save customer details first', 'error'); return; }
  currentDocType = docType;
  if (uploadModalTitle) uploadModalTitle.textContent = `Upload ${docName}`;
  if (uploadDocType) uploadDocType.textContent = docName;
  singleUploadModal.classList.add('active');
};

function previewSingleFile(file) {
  const fileType = file.type;
  const isImage = fileType.startsWith('image/');
  if (singlePreviewObjectUrl) { URL.revokeObjectURL(singlePreviewObjectUrl); singlePreviewObjectUrl = null; }
  const previewSrc = isImage ? URL.createObjectURL(file) : '';
  if (isImage) singlePreviewObjectUrl = previewSrc;
  if (singleFilePreview) {
    singleFilePreview.innerHTML = `
      <div class="single-file-preview">
        ${isImage ? `<img src="${previewSrc}" class="preview-image">` : `<i class="fas ${getFileIcon(file.name)} file-icon-large"></i>`}
        <div class="file-details">
          <div class="file-name">${file.name}</div>
          <div class="file-size">${formatFileSize(file.size)}</div>
        </div>
      </div>
    `;
  }
}

function clearSingleFilePreview() {
  if (singlePreviewObjectUrl) { URL.revokeObjectURL(singlePreviewObjectUrl); singlePreviewObjectUrl = null; }
  if (singleFilePreview) singleFilePreview.innerHTML = '';
}

function canvasToJpegBlob(canvas, quality) {
  return new Promise((resolve) => { canvas.toBlob((blob) => resolve(blob), 'image/jpeg', quality); });
}

async function compressImageTo200KB(file, targetBytes = MAX_IMAGE_UPLOAD_BYTES) {
  const dataUrl = await fileToDataURL(file);
  const img = await loadImage(dataUrl);
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { alpha: false });
  if (!ctx) throw new Error('Image processing is not supported on this device.');
  let width = img.width, height = img.height;
  const maxEdge = 1920;
  const longest = Math.max(width, height);
  if (longest > maxEdge) { const ratio = maxEdge / longest; width = Math.max(1, Math.round(width * ratio)); height = Math.max(1, Math.round(height * ratio)); }
  let bestBlob = null;
  for (let pass = 0; pass < 10; pass++) {
    canvas.width = Math.max(1, width); canvas.height = Math.max(1, height);
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    for (let quality = 0.9; quality >= 0.35; quality -= 0.07) {
      const blob = await canvasToJpegBlob(canvas, quality);
      if (!blob) continue;
      if (!bestBlob || blob.size < bestBlob.size) bestBlob = blob;
      if (blob.size <= targetBytes) {
        const baseName = (file.name || 'photo').replace(/\.[^/.]+$/, '');
        return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' });
      }
    }
    width = Math.max(360, Math.round(width * 0.85));
    height = Math.max(360, Math.round(height * 0.85));
    if (width <= 360 && height <= 360) break;
  }
  if (bestBlob && bestBlob.size <= targetBytes) {
    const baseName = (file.name || 'photo').replace(/\.[^/.]+$/, '');
    return new File([bestBlob], `${baseName}.jpg`, { type: 'image/jpeg' });
  }
  const targetMb = (targetBytes / (1024 * 1024)).toFixed(0);
  throw new Error(`Could not reduce image below ${targetMb}MB. Please retake at a lower resolution.`);
}

async function compressPdfUnderLimit(file, targetBytes = MAX_PDF_SIZE_BYTES) {
  if (!window.pdfjsLib) {
    throw new Error('PDF compression library is not available. Please refresh and try again.');
  }

  let jsPDFConstructor = null;
  if (window.jspdf && window.jspdf.jsPDF) {
    jsPDFConstructor = window.jspdf.jsPDF;
  } else if (window.jsPDF && typeof window.jsPDF === 'function') {
    jsPDFConstructor = window.jsPDF;
  } else if (typeof jspdf !== 'undefined' && jspdf.jsPDF) {
    jsPDFConstructor = jspdf.jsPDF;
  }
  if (!jsPDFConstructor) {
    throw new Error('PDF compression is not available right now. Please refresh and try again.');
  }

  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const attempts = [
    { scale: 1.15, quality: 0.7 },
    { scale: 1.0, quality: 0.58 },
    { scale: 0.9, quality: 0.5 },
    { scale: 0.8, quality: 0.42 },
    { scale: 0.7, quality: 0.36 }
  ];
  let bestBlob = null;

  for (const attempt of attempts) {
    const pdf = new jsPDFConstructor({ orientation: 'portrait', unit: 'pt', format: 'a4', compress: true });
    let firstPage = true;

    for (let pageNumber = 1; pageNumber <= pdfDoc.numPages; pageNumber++) {
      const page = await pdfDoc.getPage(pageNumber);
      const viewport = page.getViewport({ scale: attempt.scale });
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d', { alpha: false });
      if (!context) throw new Error('Canvas is not available for PDF compression.');
      canvas.width = Math.max(1, Math.floor(viewport.width));
      canvas.height = Math.max(1, Math.floor(viewport.height));
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: context, viewport }).promise;

      const imageBlob = await canvasToJpegBlob(canvas, attempt.quality);
      if (!imageBlob) throw new Error('Could not compress PDF page image.');
      const imageDataUrl = await fileToDataURL(new File([imageBlob], `page-${pageNumber}.jpg`, { type: 'image/jpeg' }));
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const ratio = Math.min(pageWidth / canvas.width, pageHeight / canvas.height);
      const drawWidth = canvas.width * ratio;
      const drawHeight = canvas.height * ratio;
      const x = (pageWidth - drawWidth) / 2;
      const y = (pageHeight - drawHeight) / 2;

      if (!firstPage) pdf.addPage();
      firstPage = false;
      pdf.addImage(imageDataUrl, 'JPEG', x, y, drawWidth, drawHeight, undefined, 'FAST');
    }

    const compressedBlob = pdf.output('blob');
    if (!bestBlob || compressedBlob.size < bestBlob.size) bestBlob = compressedBlob;
    if (compressedBlob.size <= targetBytes) {
      const baseName = (file.name || 'document').replace(/\.[^/.]+$/, '');
      return new File([compressedBlob], `${baseName}.pdf`, { type: 'application/pdf' });
    }
  }

  if (bestBlob && bestBlob.size <= targetBytes) {
    const baseName = (file.name || 'document').replace(/\.[^/.]+$/, '');
    return new File([bestBlob], `${baseName}.pdf`, { type: 'application/pdf' });
  }

  throw new Error(`Could not reduce PDF below ${getPdfUploadLimitLabel()}. Please upload a smaller scan.`);
}

async function compressFileToUploadLimit(file) {
  const isImage = Boolean(file?.type && file.type.startsWith('image/'));
  const isPdf = (file?.type === 'application/pdf') || String(file?.name || '').toLowerCase().endsWith('.pdf');

  if (isImage) {
    const compressedImage = file.size > MAX_IMAGE_UPLOAD_BYTES
      ? await compressImageTo200KB(file, MAX_IMAGE_UPLOAD_BYTES)
      : file;
    let pdfBlob = await convertImageToPdf(compressedImage);
    if (pdfBlob.size > MAX_PDF_SIZE_BYTES) {
      const converted = await convertImageToPdfUnderLimit(file);
      pdfBlob = converted.pdfBlob;
    }
    const baseName = (file.name || 'image').replace(/\.[^/.]+$/, '');
    return new File([pdfBlob], `${baseName}.pdf`, { type: 'application/pdf' });
  }

  if (isPdf) {
    if (file.size <= MAX_PDF_SIZE_BYTES) return file;
    return compressPdfUnderLimit(file, MAX_PDF_SIZE_BYTES);
  }

  throw new Error('Only PDF files and images can be compressed for upload.');
}

async function maybeCompressOversizedFile(file) {
  const isImage = Boolean(file?.type && file.type.startsWith('image/'));
  const isPdf = (file?.type === 'application/pdf') || String(file?.name || '').toLowerCase().endsWith('.pdf');
  const exceedsLimit = file.size > MAX_PDF_SIZE_BYTES;

  if (!exceedsLimit) return file;
  if (!isImage && !isPdf) {
    throw new Error('Unsupported file type. Please upload images or PDFs only.');
  }

  const proceed = await showCompressionPromptModal(file);
  if (!proceed) return null;

  showNotification(`Compressing ${file.name}...`, 'info');
  const compressed = await compressFileToUploadLimit(file);
  if (compressed.size > MAX_PDF_SIZE_BYTES) {
    throw new Error(`Compressed file is still above ${getPdfUploadLimitLabel()}. Please use a smaller file.`);
  }
  showNotification(`Compression complete: ${file.name} is now ${formatFileSize(compressed.size)}`, 'success');
  return compressed;
}

async function handleSingleFileSelection(file) {

  // FIX: Don't exit early if confirmSingleUpload is missing - still show modal
  if (!confirmSingleUpload) {
    // Silently continue
  }

  const isImage = Boolean(file?.type && file.type.startsWith('image/'));
  const isPdf = (file?.type === 'application/pdf') || (String(file?.name || '').toLowerCase().endsWith('.pdf'));

  if (confirmSingleUpload) confirmSingleUpload.disabled = true;
  clearSingleFilePreview();

  try {
    let preparedFile = file;
    if (file.size > MAX_PDF_SIZE_BYTES) {
      const compressedFile = await maybeCompressOversizedFile(file);
      if (!compressedFile) {
        currentFile = null;
        clearSingleFilePreview();
        return;
      }
      preparedFile = compressedFile;
    }
    if (isImage) {
      // Requirement: on mobile camera snap, compress only when the photo exceeds the configured image limit.
      if (preparedFile.type.startsWith('image/') && preparedFile.size > MAX_IMAGE_UPLOAD_BYTES) {
        showNotification(`Optimizing photo to ${getImageUploadLimitLabel()} for upload...`, 'info');
        preparedFile = await compressImageTo200KB(preparedFile, MAX_IMAGE_UPLOAD_BYTES);
        if (preparedFile.size > MAX_IMAGE_UPLOAD_BYTES) throw new Error(`Photo must be ${getImageUploadLimitLabel()} or less after compression.`);
      }
    }
    currentFile = preparedFile;
    previewSingleFile(preparedFile);
    if (confirmSingleUpload) confirmSingleUpload.disabled = false;
  } catch (error) {
    currentFile = null;
    showNotification(error.message || 'Could not process selected image.', 'error');
  }
}

function fileToDataURL(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = () => res(reader.result);
    reader.onerror = rej;
    reader.readAsDataURL(file);
  });
}

function loadImage(dataUrl) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = rej;
    img.src = dataUrl;
  });
}

// ==================== FIXED: IMAGE TO PDF CONVERSION ====================
async function convertImageToPdf(file) {

  try {
    // Check if jsPDF is available
    let jsPDFConstructor = null;

    // Try multiple ways to get jsPDF
    if (window.jspdf && window.jspdf.jsPDF) {
      jsPDFConstructor = window.jspdf.jsPDF;
    } else if (window.jsPDF && typeof window.jsPDF === 'function') {
      jsPDFConstructor = window.jsPDF;
    } else if (typeof jspdf !== 'undefined' && jspdf.jsPDF) {
      jsPDFConstructor = jspdf.jsPDF;
    } else {
      // Show user-friendly error
      showNotification('⚠️ Image to PDF conversion unavailable. Please upload PDF files directly.', 'error');
      throw new Error('jsPDF library not available. Please upload PDF files directly.');
    }

    const dataUrl = await fileToDataURL(file);
    const img = await loadImage(dataUrl);

    // Create PDF
    const pdf = new jsPDFConstructor({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    // Calculate image dimensions to fit on page
    let imgWidth = img.width;
    let imgHeight = img.height;

    // Calculate scaling ratio to fit image on page (with 10% margins)
    const widthRatio = (pageWidth * 0.9) / imgWidth;
    const heightRatio = (pageHeight * 0.9) / imgHeight;
    const ratio = Math.min(widthRatio, heightRatio);

    const scaledWidth = imgWidth * ratio;
    const scaledHeight = imgHeight * ratio;

    // Center on page
    const x = (pageWidth - scaledWidth) / 2;
    const y = (pageHeight - scaledHeight) / 2;

    // Add image to PDF
    pdf.addImage(dataUrl, 'JPEG', x, y, scaledWidth, scaledHeight);

    // Get PDF as blob
    const pdfBlob = pdf.output('blob');

    return pdfBlob;

  } catch (error) {
    // Check if it's a jsPDF loading error
    if (error.message && error.message.includes('jsPDF')) {
      showNotification('⚠️ Image to PDF conversion failed: jsPDF library not loaded', 'error');
    } else {
      showNotification('Image conversion failed: ' + error.message, 'error');
    }

    throw error;
  }
}

async function convertImageToPdfUnderLimit(imageFile) {
  const targets = [
    MAX_IMAGE_UPLOAD_BYTES,
    900 * 1024,
    800 * 1024,
    700 * 1024,
    600 * 1024,
    500 * 1024,
    400 * 1024,
    300 * 1024,
    250 * 1024,
    200 * 1024
  ];

  let lastError = null;
  for (const target of targets) {
    try {
      const inputFile = imageFile.size > target
        ? await compressImageTo200KB(imageFile, target)
        : imageFile;
      const pdfBlob = await convertImageToPdf(inputFile);
      if (pdfBlob.size <= MAX_PDF_SIZE_BYTES) {
        return { pdfBlob, inputFile };
      }
      lastError = new Error(`Converted PDF still larger than ${getPdfUploadLimitLabel()}`);
    } catch (e) {
      lastError = e;
    }
  }
  throw lastError || new Error(`Could not compress image to fit under ${getPdfUploadLimitLabel()} PDF limit`);
}

function findMatchingDocTypeByName(filename, usedTypes) {
  const name = filename.toLowerCase();
  const scores = DOCUMENT_TYPES.map(dt => {
    const idScore = name.includes(dt.id.replace(/_/g, ' ')) ? 10 : 0;
    const words = dt.name.toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
    let wordScore = 0;
    for (const w of words) if (name.includes(w)) wordScore += 2;
    return { id: dt.id, score: idScore + wordScore };
  }).sort((a,b) => b.score - a.score);
  for (const s of scores) {
    if (s.score <= 0) continue;
    if (usedTypes && usedTypes.includes(s.id)) continue;
    return s.id;
  }
  return null;
}

async function handleBatchFiles(files) {
  if (!files || files.length === 0) return;
  if (!customerDetailsSaved && !currentEditId) { showNotification('Please save customer details before batch uploading', 'error'); return; }
  const customerName = customerNameInput.value.trim();
  if (!customerName) { showNotification('Customer name is required', 'error'); return; }
  if (window.__uploadInProgress) return showNotification('Upload already in progress', 'error');
  window.__uploadInProgress = true;
  const total = files.length;
  if (window.showLoader) window.showLoader('Uploading batch files... 0%');
  try {
    const storage = new BackblazeStorage();
    // Track doc types already occupied (existing uploads + current batch) so two files can't map to one type.
    const usedTypes = new Set(Object.keys(currentCustomerUploads || {}).filter((id) => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
    let successCount = 0;
    let failCount = 0;
    const activeDocumentIds = new Set(DOCUMENT_TYPES.map((doc) => doc.id));

    for (let idx = 0; idx < files.length; idx++) {
      const item = files[idx];
      const file = item.file || item;
      const mappedType = item.mappedType || item.type || null;

      const isImage = Boolean(file?.type && file.type.startsWith('image/'));
      const isPdf = (file?.type === 'application/pdf') || (String(file?.name || '').toLowerCase().endsWith('.pdf'));

      let targetType = mappedType || findMatchingDocTypeByName(file.name, Array.from(usedTypes));
      if (!targetType) {
        const missing = DOCUMENT_TYPES.filter(d => d.required !== false).map(d => d.id).filter(id => !(currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
        targetType = missing.find((id) => !usedTypes.has(id)) || null;
      }
      if (!targetType) {
        const notUsed = DOCUMENT_TYPES.map(d => d.id).filter(id => !(currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
        targetType = notUsed.find((id) => !usedTypes.has(id)) || DOCUMENT_TYPES[0].id;
      }
      if (!activeDocumentIds.has(targetType)) {
        failCount++;
        continue;
      }
      // Prevent two different files from mapping to the same slot in the same batch.
      if (usedTypes.has(targetType)) {
        showNotification(`⚠️ ${file.name} could not be mapped because "${targetType}" is already selected/filled. Use batch mapping to choose another type.`, 'warning');
        failCount++;
        continue;
      }
      usedTypes.add(targetType);

      let fileToSend = file;
      try {
        if (fileToSend.size > MAX_PDF_SIZE_BYTES) {
          fileToSend = await compressFileToUploadLimit(fileToSend);
        }
        if (isImage) {
          // Check jsPDF first
          if (!window.jspdf || (!window.jspdf.jsPDF && !window.jsPDF)) {
            showNotification(`⚠️ Cannot convert ${file.name} to PDF. Please upload PDF files directly.`, 'error');
            failCount++;
            continue;
          }

          const optimizedImage = file.size > MAX_IMAGE_UPLOAD_BYTES
            ? await compressImageTo200KB(file, MAX_IMAGE_UPLOAD_BYTES)
            : file;
          let pdfBlob = await convertImageToPdf(optimizedImage);
          if (pdfBlob.size > MAX_PDF_SIZE_BYTES) {
            try {
              const converted = await convertImageToPdfUnderLimit(file);
              pdfBlob = converted.pdfBlob;
            } catch (e) {
              showNotification(`Converted ${file.name} is above ${getPdfUploadLimitLabel()}, skipped.`, 'error');
              failCount++;
              continue;
            }
          }
          const newName = file.name.replace(/\.[^/.]+$/, '') + '.pdf';
          fileToSend = new File([pdfBlob], newName, { type: 'application/pdf' });
        } else if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          // Already validated
        } else {
          showNotification(`${file.name} unsupported, only images or PDFs allowed`, 'error');
          failCount++;
          continue;
        }
      } catch (e) {
        showNotification(`Failed to convert ${file.name}`, 'error');
        failCount++;
        continue;
      }

      try {
        const result = await storage.uploadFile(fileToSend, customerName, targetType);
        setUploadedDocForType(targetType, { name: fileToSend.name, fileId: result.fileId, fileUrl: result.fileUrl, uploadedAt: await getTrustedNowIso(), localAddedAt: Date.now() });
        successCount++;
      } catch (e) {
        console.error('Batch upload failed for file', file.name, e);
        const reason = String(e?.message || 'Unknown upload error');
        showNotification(`Upload failed for ${file.name}: ${reason}`, 'error');
        failCount++;
      }

      try {
        const done = idx + 1;
        const percent = Math.round((done / total) * 100);
        if (window.showLoader) window.showLoader(`Uploading batch files... ${percent}%`);
      } catch (e) { }
    }

    if (window.renderDocumentGridUpload) {
      window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'edit');
      window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'edit');
    }
    updateSubmitButton();
    await persistDraftSilentlyIfNeeded();
    showNotification(`Batch upload complete: ${successCount} successful, ${failCount} failed`, successCount > 0 ? 'success' : 'error');
  } finally {
    hideLoader();
    window.__uploadInProgress = false;
    __batchFilesBuffer = [];
  }
}

function prepareBatchForMapping(files) {
  if (!files || files.length === 0) return;
  const alreadyUploaded = new Set(Object.keys(currentCustomerUploads || {}).filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
  const usedInBatch = new Set(Array.from(alreadyUploaded));
  __batchFilesBuffer = files.map(f => {
    const g = findMatchingDocTypeByName(f.name, Array.from(usedInBatch));
    const selected = (g && !usedInBatch.has(g)) ? g : null;
    if (selected) usedInBatch.add(selected);
    return { originalFile: f, previewName: f.name, guessedType: g, selectedType: selected };
  });
  renderBatchMappingList();
  if (batchMappingModal) batchMappingModal.classList.add('active');
}

function renderBatchMappingList() {
  if (!batchMappingList) return;
  const alreadyUploaded = new Set(Object.keys(currentCustomerUploads || {}).filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
  batchMappingList.innerHTML = __batchFilesBuffer.map((entry, idx) => {
    const options = DOCUMENT_TYPES.map(d => {
      const disabled = alreadyUploaded.has(d.id) && d.id !== entry.selectedType ? 'disabled' : '';
      const sel = d.id === entry.selectedType ? 'selected' : '';
      return `<option value="${d.id}" ${sel} ${disabled}>${d.name}</option>`;
    }).join('');
    return `
      <div class="batch-item" style="display:flex; gap:12px; align-items:center; padding:8px; border:1px solid #eef2f7; border-radius:8px;">
        <div style="width:40px; text-align:center;"><i class="fas ${getFileIcon(entry.previewName)}"></i></div>
        <div style="flex:1;">
          <div style="font-weight:600;">${entry.previewName}</div>
          <div style="font-size:12px; color:var(--text-muted);">${formatFileSize(entry.originalFile.size)}</div>
        </div>
        <div>
          <select data-idx="${idx}" class="batch-map-select">
            <option value="">-- Skip --</option>
            ${options}
          </select>
        </div>
      </div>
    `;
  }).join('');
  const selects = batchMappingList.querySelectorAll('.batch-map-select');
  selects.forEach(s => {
    const idx = parseInt(s.dataset.idx, 10);
    const sel = __batchFilesBuffer[idx].selectedType;
    if (sel) s.value = sel;
    s.addEventListener('change', (e) => { __batchFilesBuffer[idx].selectedType = e.target.value || null; refreshBatchSelectOptions(); });
  });
  refreshBatchSelectOptions();
}

function refreshBatchSelectOptions() {
  if (!batchMappingList) return;
  // Enforce uniqueness: if a doc type is chosen multiple times, keep the first and clear the rest.
  const seen = new Map(); // docType -> firstIdx
  let hadDuplicates = false;
  __batchFilesBuffer.forEach((b, idx) => {
    const t = b.selectedType;
    if (!t) return;
    if (!seen.has(t)) { seen.set(t, idx); return; }
    hadDuplicates = true;
    b.selectedType = null;
  });

  if (hadDuplicates) {
    // Sync UI to cleared selections.
    const selects = batchMappingList.querySelectorAll('.batch-map-select');
    selects.forEach(s => {
      const idx = parseInt(s.dataset.idx, 10);
      s.value = __batchFilesBuffer[idx].selectedType || '';
    });
    showNotification('Duplicate document mapping removed. Please reselect the cleared ones.', 'warning');
  }

  const chosen = new Set(__batchFilesBuffer.map(b => b.selectedType).filter(Boolean));
  const alreadyUploaded = new Set(Object.keys(currentCustomerUploads || {}).filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
  const selects = batchMappingList.querySelectorAll('.batch-map-select');
  selects.forEach(s => {
    const idx = parseInt(s.dataset.idx, 10);
    const mySelected = __batchFilesBuffer[idx].selectedType;
    Array.from(s.options).forEach(opt => {
      if (!opt.value) { opt.disabled = false; return; }
      if (alreadyUploaded.has(opt.value) && opt.value !== mySelected) { opt.disabled = true; return; }
      opt.disabled = (opt.value !== mySelected && chosen.has(opt.value));
    });
  });
}

async function uploadSingleDocument() {
  if (!assertWritable('Document upload')) return;
  if (!currentFile || !currentDocType) return;
  if (!customerDetailsSaved && !currentEditId) { showNotification('Please save customer details first', 'error'); return; }
  if (currentFile.size > MAX_PDF_SIZE_BYTES) {
    try {
      const compressed = await maybeCompressOversizedFile(currentFile);
      if (!compressed) return;
      currentFile = compressed;
      previewSingleFile(currentFile);
    } catch (error) {
      showNotification(error.message || 'Could not compress the selected file.', 'error');
      return;
    }
  }
  if (window.__uploadInProgress) { return; }
  window.__uploadInProgress = true;
  if (confirmSingleUpload) {
    confirmSingleUpload.disabled = true;
    confirmSingleUpload.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Uploading...';
  }
  try {
    const storage = new BackblazeStorage();
    const customerName = customerNameInput.value.trim();
    if (!customerName) throw new Error('Please enter customer name first');
    let fileToSend = currentFile;
    if (currentFile.type.startsWith('image/')) {
      try {
        let pdfBlob = null;
        try {
          // Try a normal conversion first (fast path).
          const baseImage = currentFile.size > MAX_IMAGE_UPLOAD_BYTES
            ? await compressImageTo200KB(currentFile, MAX_IMAGE_UPLOAD_BYTES)
            : currentFile;
          pdfBlob = await convertImageToPdf(baseImage);
        } catch (e) {
          throw e;
        }

        if (pdfBlob.size > MAX_PDF_SIZE_BYTES) {
          const ok = confirm(`Converted PDF is larger than ${getPdfUploadLimitLabel()}.\n\nTry compressing the photo further and retry?`);
          if (!ok) return;
          const converted = await convertImageToPdfUnderLimit(currentFile);
          pdfBlob = converted.pdfBlob;
        }
        const newName = currentFile.name.replace(/\.[^/.]+$/, '') + '.pdf';
        fileToSend = new File([pdfBlob], newName, { type: 'application/pdf' });
      } catch (e) {
        showNotification('Image conversion failed: ' + e.message, 'error');
        return;
      }
    } else if (currentFile.type === 'application/pdf' || currentFile.name.toLowerCase().endsWith('.pdf')) {
      // Already validated
    } else {
      showNotification('Unsupported file type. Please upload images or PDFs.', 'error');
      return;
    }
    const result = await storage.uploadFile(fileToSend, customerName, currentDocType);
    setUploadedDocForType(currentDocType, { name: currentFile.name, fileId: result.fileId, fileUrl: result.fileUrl, uploadedAt: await getTrustedNowIso(), localAddedAt: Date.now() });
    renderDocumentGrid(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'upload');
    renderDocumentGrid(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'upload');
    updateSubmitButton();
    await persistDraftSilentlyIfNeeded();
    showNotification('✅ Document uploaded successfully!', 'success');
    closeModal(singleUploadModal);
  } catch (error) {
    console.error('Single document upload failed', error);
    showNotification('Upload failed: ' + error.message, 'error');
  } finally {
    if (confirmSingleUpload) {
      confirmSingleUpload.disabled = false;
      confirmSingleUpload.innerHTML = 'Upload Document';
    }
    window.__uploadInProgress = false;
  }
}

// ==================== VIEW DOCUMENT ====================
window.viewDocument = (docType) => {
  const docs = currentCustomerUploads[docType];
  if (!docs || docs.length === 0) return;
  const lastDoc = docs[docs.length - 1];
  window.open(lastDoc.fileUrl, '_blank');
};

// ==================== VIEW SUBMISSION DOCUMENTS ====================
window.viewSubmissionDocs = (submissionId) => {
  const sub = allSubmissions.find(s => s.id === submissionId);
  const effectiveDocs = getEffectiveSubmissionDocuments(sub);
  if (!sub || effectiveDocs.length === 0) { showNotification('No documents available', 'error'); return; }
  const firstDoc = effectiveDocs[0];
  const docTypeLabel = DOCUMENT_TYPES.find(t => t.id === firstDoc.documentType)?.name || firstDoc.documentType || 'Document';
  if (viewerModal && viewerFileName && documentViewer) {
    viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel}`;
    documentViewer.src = firstDoc.fileUrl?.trim();
    viewerModal.classList.add('active');
  }
  if (effectiveDocs.length > 1) {
    let currentIndex = 0;
    const showDoc = (index) => {
      const doc = effectiveDocs[index];
      const docTypeLabel = DOCUMENT_TYPES.find(t => t.id === doc.documentType)?.name || doc.documentType || 'Document';
      viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel} (${index + 1}/${effectiveDocs.length})`;
      documentViewer.src = doc.fileUrl?.trim();
    };
    const addViewerNav = () => {
      const viewerHeader = viewerModal ? viewerModal.querySelector('.viewer-header') : document.querySelector('.viewer-header');
      if (!viewerHeader) return;
      const existingNav = viewerHeader.querySelector('.viewer-nav');
      if (existingNav) existingNav.remove();
      const nav = document.createElement('div');
      nav.className = 'viewer-nav';
      nav.style.cssText = 'display: flex; gap: 10px; align-items: center;';
      nav.innerHTML = `
        <button id="prevDoc" class="action-btn" ${currentIndex === 0 ? 'disabled' : ''}><i class="fas fa-chevron-left"></i> Prev</button>
        <span id="docCounter" style="font-size: 14px; color: #666;">${currentIndex + 1}/${effectiveDocs.length}</span>
        <button id="nextDoc" class="action-btn" ${currentIndex === effectiveDocs.length - 1 ? 'disabled' : ''}><i class="fas fa-chevron-right"></i> Next</button>
      `;
      const closeBtn = viewerHeader.querySelector('#closeViewer') || document.getElementById('closeViewer');
      viewerHeader.insertBefore(nav, closeBtn);
      document.getElementById('prevDoc').onclick = () => { if (currentIndex > 0) { currentIndex--; showDoc(currentIndex); addViewerNav(); } };
      document.getElementById('nextDoc').onclick = () => { if (currentIndex < effectiveDocs.length - 1) { currentIndex++; showDoc(currentIndex); addViewerNav(); } };
    };
    addViewerNav();
  }
};

// ==================== SHOW APPLICATION TRACKING ====================
window.showApplicationTrack = async (submissionId) => {
  let sub = allSubmissions.find(s => s.id === submissionId);
  if (!sub && submissionId) {
    try {
      const subSnap = await getDoc(doc(db, 'submissions', submissionId));
      if (subSnap.exists()) {
        sub = { id: subSnap.id, ...(subSnap.data() || {}) };
      }
    } catch (error) {
      console.warn('Could not fetch application for tracking', error);
    }
  }
  if (!sub) {
    showNotification('Application not found', 'error');
    return;
  }
  if (String(sub.status || '').toLowerCase() === 'draft') {
    showNotification('Draft applications are managed from the Draft tab.', 'info');
    return;
  }
  try {
    const stage = await getApplicationStage(sub);
    const trackModal = document.createElement('div');
    trackModal.className = 'modal';
    trackModal.id = 'trackModal';
    trackModal.style.cssText = 'display: flex !important; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 999999; align-items: center; justify-content: center;';
    trackModal.innerHTML = `
      <div class="modal-content" style="max-width: 500px;">
        <div class="modal-header">
          <h2><i class="fas fa-map-marker-alt"></i> Application Tracking</h2>
          <button class="close-btn" onclick="this.closest('.modal').remove()">&times;</button>
        </div>
        <div class="modal-body" style="padding: 30px;">
          <div style="text-align: center; margin-bottom: 30px;">
            <i class="fas fa-file-alt" style="font-size: 60px; color: var(--cm-primary);"></i>
            <h3 style="margin: 15px 0 5px;">${sub.customerName}</h3>
            <p style="color: var(--gray-500);">Application ID: ${sub.id.substring(0, 8)}...</p>
          </div>
          <div style="background: #f8fafc; border-radius: 12px; padding: 20px;">
            <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 15px;">
              <div style="width: 40px; height: 40px; background: ${getStageColor(stage)}; border-radius: 50%; display: flex; align-items: center; justify-content: center;">
                <i class="fas ${getStageIcon(stage)}" style="color: white;"></i>
              </div>
              <div>
                <div style="font-size: 12px; color: var(--gray-500);">Current Stage</div>
                <div style="font-size: 18px; font-weight: 600; color: var(--gray-800);">${stage}</div>
              </div>
            </div>
            <div style="border-top: 1px solid var(--gray-200); padding-top: 15px;">
              <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                <span style="color: var(--gray-600);">Status:</span>
                <span class="status-badge status-${sub.status}">${sub.status}</span>
              </div>
              <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                <span style="color: var(--gray-600);">Uploaded:</span>
                <span>${safeFormatDate(sub.uploadedAt)}</span>
              </div>
              ${sub.assignedTo ? `<div style="display: flex; justify-content: space-between; margin-bottom: 10px;"><span style="color: var(--gray-600);">Assigned To:</span><span>${sub.assignedTo}</span></div>` : ''}
              ${sub.reviewedAt ? `<div style="display: flex; justify-content: space-between; margin-bottom: 10px;"><span style="color: var(--gray-600);">Reviewed:</span><span>${safeFormatDate(sub.reviewedAt)}</span></div>` : ''}
              ${sub.comment ? `<div style="margin-top: 15px;"><span style="color: var(--gray-600);">Comment:</span><p style="background: white; padding: 10px; border-radius: 8px; margin-top: 5px;">${sub.comment}</p></div>` : ''}
            </div>
          </div>
          <div style="margin-top: 25px; text-align: center;">
            <button class="action-btn" onclick="this.closest('.modal').remove()" style="padding: 10px 30px;">Close</button>
          </div>
        </div>
      </div>
    `;
    const target = document.body || document.documentElement;
    if (target) target.appendChild(trackModal);
    trackModal.addEventListener('click', (e) => { if (e.target === trackModal) trackModal.remove(); });
  } catch (error) {
    showNotification('Could not load tracking information', 'error');
  }
};

function getStageColor(stage) {
  if (stage.includes('Reviewer')) return '#3b82f6';
  if (stage.includes('Approved')) return '#10b981';
  if (stage.includes('Rejected')) return '#ef4444';
  if (stage.includes('RSA')) return '#8b5cf6';
  return '#f59e0b';
}
function getStageIcon(stage) {
  if (stage.includes('Reviewer')) return 'fa-user-check';
  if (stage.includes('Approved')) return 'fa-check-circle';
  if (stage.includes('Rejected')) return 'fa-times-circle';
  if (stage.includes('RSA')) return 'fa-building';
  return 'fa-clock';
}

// ==================== SUBMIT CUSTOMER ====================
function updateSubmitButton() {
  if (!submitCustomerBtn) return;
  const customerName = customerNameInput.value.trim();
  const allowWithoutDocuments = canCurrentUserSubmitWithoutDocuments();
  const hasAnyDoc = Object.keys(currentCustomerUploads || {}).some(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0);
  if (currentEditId) {
    uploadedCountSpan.textContent = Object.keys(currentCustomerUploads || {}).filter((id) => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0).length;
    submitCustomerBtn.disabled = !(customerName && (hasAnyDoc || allowWithoutDocuments));
    syncUploadRequirementUi();
    return;
  }
  const requiredDocs = DOCUMENT_TYPES.filter(d => d.required !== false).map(d => d.id);
  const uploadedRequired = requiredDocs.filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0).length;
  uploadedCountSpan.textContent = Object.keys(currentCustomerUploads || {}).filter((id) => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0).length;
  submitCustomerBtn.disabled = !(customerName && customerDetailsSaved && (allowWithoutDocuments || uploadedRequired === requiredDocs.length));
  syncUploadRequirementUi();
}

const REQUIRED_SUBMISSION_FIELD_KEYS = {
  customerName: ['name', 'customerName'],
  customerDob: ['dob', 'dateOfBirth', 'customerDob'],
  customerEmail: ['email', 'customerEmail'],
  customerPhone: ['phone', 'customerPhone'],
  customerNIN: ['nin', 'customerNIN'],
  customerAddress: ['address', 'customerAddress'],
  accountNo: ['accountNo'],
  employer: ['employer'],
  originatingTP: ['originatingTP'],
  mortgageLoanApplicationFormDate: ['mortgageLoanApplicationFormDate'],
  pfa: ['pfa', 'pfaName'],
  penNo: ['penNo'],
  rsaStatementDate: ['rsaStatementDate'],
  rsaBalance: ['rsaBalance'],
  propertyType: ['propertyType'],
  propertyValue: ['propertyValue'],
  facilityFee: ['facilityFee'],
  loanAmount: ['loanAmount']
};

function getStoredSubmissionFieldValue(submission = {}, fieldId = '') {
  const keys = REQUIRED_SUBMISSION_FIELD_KEYS[fieldId] || [fieldId];
  return getSubmissionDetailValue(submission, keys, '');
}

function getDerivedSubmissionFieldValue(fieldId = '', submission = getCurrentEditableSubmission()) {
  const rsaBalance = parseMoney(getFormValue('rsaBalance') || getStoredSubmissionFieldValue(submission, 'rsaBalance'));
  if (!rsaBalance) return '';
  const rule = determinePropertyByRsa(rsaBalance);
  if (!rule) return '';
  if (fieldId === 'propertyType') return String(rule.name || '');
  if (fieldId === 'propertyValue') return String(rule.value || '');
  if (fieldId === 'facilityFee') return String(rule.fee || '');
  if (fieldId === 'loanAmount') {
    return String(roundUpToNearestThousand(Number(rule.value || 0) - calculateRoundedRsa25(rsaBalance)));
  }
  return '';
}

function setFieldValueIfEmpty(fieldId = '', value = '') {
  const el = document.getElementById(fieldId);
  if (!el || String(el.value || '').trim() || String(value || '').trim() === '') return;
  el.value = value;
}

function syncMoneyDisplayField(fieldId = '', displayId = '') {
  const value = getFormValue(fieldId);
  const displayEl = document.getElementById(displayId);
  if (displayEl && value) displayEl.textContent = formatCurrency(value);
}

function hydrateEditableSubmissionRequiredFields() {
  if (!currentEditId && !currentDraftId) return;
  const submission = getCurrentEditableSubmission();
  if (!submission) return;

  Object.keys(REQUIRED_SUBMISSION_FIELD_KEYS).forEach((fieldId) => {
    const storedValue = getStoredSubmissionFieldValue(submission, fieldId);
    setFieldValueIfEmpty(fieldId, storedValue);
  });

  ['propertyType', 'propertyValue', 'facilityFee', 'loanAmount'].forEach((fieldId) => {
    setFieldValueIfEmpty(fieldId, getDerivedSubmissionFieldValue(fieldId, submission));
  });

  const rsaBalance = parseMoney(getFormValue('rsaBalance') || getStoredSubmissionFieldValue(submission, 'rsaBalance'));
  if (rsaBalance && !String(getFormValue('rsa25Percent') || '').trim()) {
    const rsa25Rounded = calculateRoundedRsa25(rsaBalance);
    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) rsa25PercentEl.value = rsa25Rounded;
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) rsa25FormattedEl.textContent = formatCurrency(rsa25Rounded);
  }

  syncMoneyDisplayField('propertyValue', 'propertyValueFormatted');
  syncMoneyDisplayField('facilityFee', 'facilityFeeFormatted');
  syncMoneyDisplayField('loanAmount', 'loanAmountFormatted');
}

function getRequiredSubmissionFieldValue(fieldId = '') {
  const formValue = String(getFormValue(fieldId) || '').trim();
  if (formValue) return formValue;
  const submission = getCurrentEditableSubmission();
  const storedValue = submission ? String(getStoredSubmissionFieldValue(submission, fieldId) || '').trim() : '';
  if (storedValue) return storedValue;
  return String(getDerivedSubmissionFieldValue(fieldId, submission) || '').trim();
}

function getMissingRequiredSubmissionFields() {
  hydrateEditableSubmissionRequiredFields();
  const allowOptionalCustomerFields = isCurrentUserUploaderLevel2();
  const requiredFields = [
    { id: 'customerName', label: 'Customer Name' }, { id: 'customerDob', label: 'Date of Birth' },
    { id: 'customerEmail', label: 'Email' }, { id: 'customerPhone', label: 'Phone' },
    { id: 'customerNIN', label: 'NIN' }, { id: 'customerAddress', label: 'Address' },
    { id: 'accountNo', label: 'Account Number' }, { id: 'employer', label: 'Employer' },
    { id: 'originatingTP', label: 'Originating Transfer Pin' }, { id: 'mortgageLoanApplicationFormDate', label: 'Mortgage Loan Application Form Date' }, { id: 'pfa', label: 'PFA' },
    { id: 'penNo', label: 'PEN Number' }, { id: 'rsaStatementDate', label: 'RSA Statement Date' },
    { id: 'rsaBalance', label: 'RSA Balance' }, { id: 'propertyType', label: 'Property Type' },
    { id: 'propertyValue', label: 'Property Value' }, { id: 'loanAmount', label: 'Loan Amount' },
  ];
  const missingLabels = [], invalidLabels = [];
  requiredFields.forEach(field => {
    const el = document.getElementById(field.id);
    if (!el) return;
    const value = getRequiredSubmissionFieldValue(field.id);
    const isRequiredForUser = !allowOptionalCustomerFields || field.id === 'customerName';
    if (!value) {
      if (isRequiredForUser) missingLabels.push(field.label);
      return;
    }
    if (field.id === 'accountNo' && !/^\d{10}$/.test(value)) invalidLabels.push('Account Number (must be exactly 10 digits)');
    if (field.id === 'customerPhone' && !/^\d{11}$/.test(value)) invalidLabels.push('Phone (must be exactly 11 digits)');
    if (field.id === 'customerNIN' && !/^\d{11}$/.test(value)) invalidLabels.push('NIN (must be exactly 11 digits)');
  });
  return invalidLabels.length > 0 ? [...missingLabels, ...invalidLabels] : missingLabels;
}

async function submitCustomer() {
  if (submissionInProgress) {
    showNotification('Submission is already in progress. Please wait.', 'info');
    return;
  }
  if (!assertWritable('Submission')) return;
  const customerName = customerNameInput.value.trim();
  const allowWithoutDocuments = canCurrentUserSubmitWithoutDocuments();
  const allowOptionalCustomerFields = isCurrentUserUploaderLevel2();
  if (!customerName) return;
  if (!customerDetailsSaved) { showNotification('Please save customer details before submitting', 'error'); return; }
  submissionInProgress = true;
  submitCustomerBtn.disabled = true;
  submitCustomerBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Submitting...';
  try {
    const rolePermissions = (await getSystemSettings(db, { force: true })).rolePermissions || {};
    const currentRole = String(currentUserProfile?.role || 'uploader').trim().toLowerCase();
    if (currentRole === 'uploader' && rolePermissions.uploaderCanUpload === false) {
      showNotification('Uploader submissions are currently disabled by Super Admin.', 'error');
      return;
    }
    const missingFields = getMissingRequiredSubmissionFields();
    if (missingFields.length > 0) {
      submitCustomerBtn.disabled = false;
      syncUploadRequirementUi();
      return showNotification('All fields are compulsory. Missing or invalid: ' + missingFields.join(', '), 'error');
    }
    if (!(await resolveCustomerAccountName({ silent: false }))) {
      submitCustomerBtn.disabled = false;
      syncUploadRequirementUi();
      return;
    }
    const duplicateExcludeId = currentEditId || currentDraftId || '';
    const duplicateCheckDetails = collectDraftableCustomerDetails();
    duplicateCheckDetails.name = customerName;
    const duplicate = await validateCustomerDuplicateContact({
      phone: duplicateCheckDetails.phone,
      accountNo: duplicateCheckDetails.accountNo,
      nin: duplicateCheckDetails.nin,
      penNo: duplicateCheckDetails.penNo,
      customerName,
      excludeSubmissionId: duplicateExcludeId
    });
    if (duplicate) {
      if (duplicate.message.toLowerCase().includes('pen')) showPenNoError(formatPenNoDuplicateMessage(duplicate));
      showNotification(duplicate.message, 'error');
      submitCustomerBtn.disabled = false;
      syncUploadRequirementUi();
      return;
    }
    if (currentEditId) {
      const hasAnyDoc = Object.keys(currentCustomerUploads || {}).some(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0);
      if (!hasAnyDoc && !allowWithoutDocuments) {
        submitCustomerBtn.disabled = false;
        syncUploadRequirementUi();
        return showNotification('Please upload at least one document before submitting fix.', 'error');
      }
      const documents = collectSubmissionDocuments();
      const agentPayload = getSelectedAgentPayload();
      const customerDetails = {
        ...collectDraftableCustomerDetails(),
        name: customerName
      };
      if (!customerDetails.houseNumber && customerDetails.propertyType) {
        customerDetails.houseNumber = await reserveHouseNumber(customerDetails.propertyType);
        const houseNumberEl = document.getElementById('houseNumber');
        if (houseNumberEl && customerDetails.houseNumber) {
          houseNumberEl.value = customerDetails.houseNumber;
        }
      }
      if (!customerDetails.houseNumber && !allowOptionalCustomerFields) {
        submitCustomerBtn.disabled = false;
        syncUploadRequirementUi();
        return showNotification('Unable to generate house number for this property type.', 'error');
      }
      const submissionRef = doc(db, 'submissions', currentEditId);
      const existingSub = allSubmissions.find((s) => s.id === currentEditId) || {};
      const reviewerCandidate = String(existingSub.reviewedBy || existingSub.assignedTo || '').trim();
      const reviewerToReassign = await isActiveUserWithRole(reviewerCandidate, ['reviewer']) ? reviewerCandidate : '';
      const latestRejectedStage = String(existingSub.latestRejectedStage || '').trim().toLowerCase();
      const rsaRejectFlow = String(existingSub.status || '').toLowerCase() === 'rejected_by_rsa' || latestRejectedStage === 'rsa';
      const rsaOfficerCandidate = String(existingSub.latestRejectedBy || existingSub.assignedToRSA || '').trim();
      const rsaToReassign = await isActiveUserWithRole(rsaOfficerCandidate, ['rsa']) ? rsaOfficerCandidate : String(existingSub.assignedToRSA || '').trim();
      const nextFixCount = Number(existingSub.fixCount || 0) + 1;
      const rejectionHistory = Array.isArray(existingSub.rejectionHistory)
        ? existingSub.rejectionHistory.filter((entry) => String(entry || '').trim())
        : [];
      const previousReason = String(existingSub.latestRejectionReason || existingSub.previousRejectionReason || existingSub.comment || '').trim();
      const previousRejectedBy = String(existingSub.latestRejectedBy || existingSub.previousRejectedBy || existingSub.reviewedBy || '').trim();
      const previousRejectedAt = existingSub.latestRejectedAt || existingSub.previousRejectedAt || existingSub.reviewedAt || null;
      const preservedCommissionRate = resolveSubmissionCommissionRate(existingSub);
      const correctionPayload = {
        customerName, customerDetails, status: rsaRejectFlow ? 'processing_to_pfa' : 'pending', documents,
        documentTypes: getUploadedDocumentTypes(), reuploadedAt: serverTimestamp(),
        houseNumber: customerDetails.houseNumber || '',
        penNoNormalized: normalizePenNumber(customerDetails.penNo),
        agentId: agentPayload.agentId,
        agentName: agentPayload.agentName,
        agentContactNumber: agentPayload.agentContactNumber,
        agentAccountNumber: agentPayload.agentAccountNumber,
        agentAccountBank: agentPayload.agentAccountBank,
        fixSubmitted: true, fixLocked: false, fixSubmittedAt: serverTimestamp(), fixCount: nextFixCount,
        assignedTo: rsaRejectFlow ? '' : (reviewerToReassign || existingSub.assignedTo || ''),
        assignedToRSA: rsaRejectFlow ? (rsaToReassign || existingSub.assignedToRSA || '') : (existingSub.assignedToRSA || ''),
        rsaAssignedAt: rsaRejectFlow ? serverTimestamp() : (existingSub.rsaAssignedAt || null),
        reviewedAt: rsaRejectFlow ? (existingSub.reviewedAt || null) : null,
        reviewerDecision: rsaRejectFlow ? String(existingSub.reviewerDecision || '').trim() : '',
        reviewerDecisionBy: rsaRejectFlow ? String(existingSub.reviewerDecisionBy || '').trim() : '',
        reviewerDecisionAt: rsaRejectFlow ? (existingSub.reviewerDecisionAt || null) : null,
        rsaReady: rsaRejectFlow ? true : Boolean(existingSub.rsaReady),
        comment: previousReason,
        rejectionHistory,
        previousRejectionReason: previousReason,
        previousRejectedBy,
        previousRejectedAt,
        resubmittedAfterRejection: true,
        latestRejectedStage: '',
        commissionRate: preservedCommissionRate,
        commissionRatePercent: Number((preservedCommissionRate * 100).toFixed(4)),
        commissionRateLabel: formatCommissionRateLabel(preservedCommissionRate)
      };
      await saveSubmissionWithUniqueLocks(submissionRef, correctionPayload, { mode: 'update' });
      if (!rsaRejectFlow && !reviewerToReassign) {
        await assignRoundRobin(submissionRef);
      }
      notifyStatusChangePush({
        currentUser,
        submissionId: currentEditId,
        customerName,
        newStatus: 'pending',
        statusLabel: 'Pending Review',
        actionLabel: 'Application Re-Submitted',
        message: `Application for ${customerName} was re-submitted and is back in pending review.`
      }).catch(() => {});
      showNotification(rsaRejectFlow
        ? '✅ Fix submitted and returned directly to RSA.'
        : (reviewerToReassign
            ? '✅ Fix submitted and reassigned to the same reviewer for another review.'
            : '✅ Fix submitted successfully!'), 'success');
      closeModal(uploadModal);
      currentEditId = null;
      updateSubmitButton();
      return;
    }
    const requiredDocs = DOCUMENT_TYPES.filter(d => d.required !== false).map(d => d.id);
    const missing = requiredDocs.filter(id => !(currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
    if (!allowWithoutDocuments && missing.length > 0) {
      submitCustomerBtn.disabled = false;
      syncUploadRequirementUi();
      return showNotification('Please upload required documents: ' + missing.join(', '), 'error');
    }
    const documents = collectSubmissionDocuments();
    const customerDetails = collectDraftableCustomerDetails();
    customerDetails.name = customerName;
    const agentPayload = getSelectedAgentPayload();
    const proceed = await showNewSubmissionConfirmModal({
      customerName,
      agentPayload,
      documentsCount: documents.length
    });
    if (!proceed) {
      submitCustomerBtn.disabled = false;
      syncUploadRequirementUi();
      return;
    }
    const allocatedHouseNumber = customerDetails.houseNumber || (customerDetails.propertyType ? await reserveHouseNumber(customerDetails.propertyType) : '');
    if (!allocatedHouseNumber && !allowOptionalCustomerFields) {
      submitCustomerBtn.disabled = false;
      syncUploadRequirementUi();
      return showNotification('Unable to generate house number for this property type.', 'error');
    }
    customerDetails.houseNumber = allocatedHouseNumber || customerDetails.houseNumber || '';
    const houseNumberEl = document.getElementById('houseNumber');
    if (houseNumberEl && allocatedHouseNumber) {
      houseNumberEl.value = allocatedHouseNumber;
    }
    const uploaderEmail = normalizeEmail(currentUser?.email);
    const commissionSettings = await getCommissionSettings(db);
    const commissionFields = buildSubmissionCommissionFields(commissionSettings);
    const submissionPayload = {
      customerName,
      customerDetails,
      uploadedBy: uploaderEmail,
      status: 'pending',
      comment: '',
      documents,
      documentTypes: getUploadedDocumentTypes(),
      houseNumber: customerDetails.houseNumber || '',
      penNoNormalized: normalizePenNumber(customerDetails.penNo),
      agentId: agentPayload.agentId,
      agentName: agentPayload.agentName,
      agentContactNumber: agentPayload.agentContactNumber,
      agentAccountNumber: agentPayload.agentAccountNumber,
      agentAccountBank: agentPayload.agentAccountBank,
      ...commissionFields,
      submittedAt: serverTimestamp()
    };
    let subRef;
    if (currentDraftId) {
      subRef = doc(db, 'submissions', currentDraftId);
      await saveSubmissionWithUniqueLocks(subRef, {
        ...submissionPayload,
        uploadedAt: serverTimestamp()
      }, { mode: 'update' });
    } else {
      subRef = doc(collection(db, 'submissions'));
      await saveSubmissionWithUniqueLocks(subRef, {
        ...submissionPayload,
        uploadedAt: serverTimestamp()
      }, { mode: 'set' });
    }
    const routingRule = await getUploaderRoutingRule(uploaderEmail);
    const systemSettings = await getSystemSettings(db);
    const effectiveRouteMode = routingRule?.routeMode || String(systemSettings.routingPolicies?.defaultRouteMode || 'normal').trim().toLowerCase();
    if (effectiveRouteMode === 'skip_reviewer') {
      const assignedRsa = await assignDirectToRSA(subRef, uploaderEmail).catch(() => '');
      showNotification(assignedRsa ? `✅ Submitted and routed directly to RSA: ${assignedRsa}` : '✅ Submitted and routed directly to RSA queue.', 'success');
      notifyStatusChangePush({
        currentUser,
        submissionId: subRef.id,
        customerName,
        newStatus: 'processing_to_pfa',
        statusLabel: 'Processing to PFA',
        actionLabel: 'Direct RSA Routing',
        message: `Application for ${customerName} skipped reviewer and was routed directly to RSA.`
      }).catch(() => {});
    } else {
      const assignedEmail = await assignRoundRobin(subRef).catch(err => { return null; });
      let assignmentEmailFailed = false;
      if (assignedEmail) {
        const emailResult = await queueViewerAssignmentEmail({
          submissionId: subRef.id,
          viewerEmail: assignedEmail,
          customerName,
          uploaderEmail: currentUser?.email || '',
          uploaderName: currentUserProfile?.fullName || currentUser?.displayName || ''
        }).catch((emailErr) => {
          return { queued: true, sent: false, reason: 'send-failed' };
        });

        if (emailResult?.sent === false) {
          assignmentEmailFailed = true;
        }
      }
      if (assignmentEmailFailed) {
        showNotification(`✅ Submitted and assigned to ${assignedEmail}, but email alert failed.`, 'warning');
      } else {
        showNotification(assignedEmail ? `✅ Submitted – assigned to ${assignedEmail}` : '✅ Customer documents submitted successfully!', 'success');
      }
      notifyStatusChangePush({
        currentUser,
        submissionId: subRef.id,
        customerName,
        newStatus: 'pending',
        statusLabel: 'Pending Review',
        actionLabel: 'New Submission',
        message: `A new application for ${customerName} has been submitted for review.`
      }).catch(() => {});
    }
    closeModal(uploadModal);
    currentDraftId = null;
  } catch (error) {
    if (error?.code === 'duplicate-submission-lock') {
      showNotification(error.message || 'Duplicate application blocked.', 'error');
    } else {
      showNotification('Submission failed: ' + error.message, 'error');
    }
  } finally {
    submissionInProgress = false;
    syncUploadRequirementUi();
    updateSubmitButton();
  }
}

// ==================== EDIT FUNCTIONS ====================
window.openEditModal = async (id) => {
  const sub = allSubmissions.find(s => s.id === id);
  if (!sub) return;
  currentEditId = id;
  currentDraftId = null;
  currentHouseNumberReserved = true;
  setUploadModalHeading('fix');
  syncUploadModalControlState('fix');
  await loadApprovedAgents();
  try {
    applyDraftFormValues(sub);
    hydrateEditableSubmissionRequiredFields();
    syncCurrentSubmissionAgentFallbackFromSubmission(sub);
    populateApprovedAgentSelect();
    if (customerAgentSelect) {
      customerAgentSelect.value = getStoredAgentSelectionValue(sub);
    }
    syncOriginatingTpField();
    if (!document.getElementById('houseNumber')?.value?.trim() && document.getElementById('propertyType')?.value?.trim()) {
      await refreshHouseNumberPreview();
    }
    customerDetailsSaved = true;
  } catch (e) { }
  const existingUploads = {};
  (sub.documents || []).forEach(doc => {
    if (!existingUploads[doc.documentType]) existingUploads[doc.documentType] = [];
    existingUploads[doc.documentType].push(doc);
  });
  currentCustomerUploads = existingUploads;
  if (batchUploadBtn) batchUploadBtn.disabled = false;
  if (window.renderDocumentGridUpload) {
    window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'edit');
    window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'edit');
  }
  updateSubmitButton();
  syncUploadRequirementUi();
  if (accountNoInput) accountNoInput.focus();
  uploadModal.classList.add('active');
};

async function submitEdit() {
  if (!assertWritable('Document re-upload')) return;
  if (!currentEditId) return;
  submitEditBtn.disabled = true;
  submitEditBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Updating...';
  try {
    const documents = collectSubmissionDocuments();
    if (documents.length > 0 || canCurrentUserSubmitWithoutDocuments()) {
      const submissionRef = doc(db, 'submissions', currentEditId);
      const existingSub = allSubmissions.find((s) => s.id === currentEditId) || {};
      const reviewerCandidate = String(existingSub.reviewedBy || existingSub.assignedTo || '').trim();
      const reviewerToReassign = await isActiveUserWithRole(reviewerCandidate, ['reviewer']) ? reviewerCandidate : '';
      const latestRejectedStage = String(existingSub.latestRejectedStage || '').trim().toLowerCase();
      const rsaRejectFlow = String(existingSub.status || '').toLowerCase() === 'rejected_by_rsa' || latestRejectedStage === 'rsa';
      const rsaOfficerCandidate = String(existingSub.latestRejectedBy || existingSub.assignedToRSA || '').trim();
      const rsaToReassign = await isActiveUserWithRole(rsaOfficerCandidate, ['rsa']) ? rsaOfficerCandidate : String(existingSub.assignedToRSA || '').trim();
      const nextFixCount = Number(existingSub.fixCount || 0) + 1;
      const rejectionHistory = Array.isArray(existingSub.rejectionHistory)
        ? existingSub.rejectionHistory.filter((entry) => String(entry || '').trim())
        : [];
      const previousReason = String(existingSub.latestRejectionReason || existingSub.previousRejectionReason || existingSub.comment || '').trim();
      const previousRejectedBy = String(existingSub.latestRejectedBy || existingSub.previousRejectedBy || existingSub.reviewedBy || '').trim();
      const previousRejectedAt = existingSub.latestRejectedAt || existingSub.previousRejectedAt || existingSub.reviewedAt || null;
      const preservedCommissionRate = resolveSubmissionCommissionRate(existingSub);
      await updateDoc(submissionRef, {
        status: rsaRejectFlow ? 'processing_to_pfa' : 'pending',
        documents,
        documentTypes: getUploadedDocumentTypes(),
        reuploadedAt: serverTimestamp(),
        fixSubmitted: true,
        fixLocked: false,
        fixSubmittedAt: serverTimestamp(),
        fixCount: nextFixCount,
        assignedTo: rsaRejectFlow ? '' : (reviewerToReassign || existingSub.assignedTo || ''),
        assignedToRSA: rsaRejectFlow ? (rsaToReassign || existingSub.assignedToRSA || '') : (existingSub.assignedToRSA || ''),
        rsaAssignedAt: rsaRejectFlow ? serverTimestamp() : (existingSub.rsaAssignedAt || null),
        reviewedAt: rsaRejectFlow ? (existingSub.reviewedAt || null) : null,
        reviewerDecision: rsaRejectFlow ? String(existingSub.reviewerDecision || '').trim() : '',
        reviewerDecisionBy: rsaRejectFlow ? String(existingSub.reviewerDecisionBy || '').trim() : '',
        reviewerDecisionAt: rsaRejectFlow ? (existingSub.reviewerDecisionAt || null) : null,
        rsaReady: rsaRejectFlow ? true : Boolean(existingSub.rsaReady),
        comment: previousReason,
        rejectionHistory,
        previousRejectionReason: previousReason,
        previousRejectedBy,
        previousRejectedAt,
        resubmittedAfterRejection: true,
        latestRejectedStage: '',
        commissionRate: preservedCommissionRate,
        commissionRatePercent: Number((preservedCommissionRate * 100).toFixed(4)),
        commissionRateLabel: formatCommissionRateLabel(preservedCommissionRate)
      });
      if (!rsaRejectFlow && !reviewerToReassign) {
        await assignRoundRobin(submissionRef);
      }
      notifyStatusChangePush({
        currentUser,
        submissionId: currentEditId,
        customerName: existingSub.customerName || '',
        newStatus: 'pending',
        statusLabel: 'Pending Review',
        actionLabel: 'Application Re-Submitted',
        message: `Application for ${existingSub.customerName || 'this customer'} was re-submitted and is back in pending review.`
      }).catch(() => {});
    }
    showNotification(
      (String((allSubmissions.find((s) => s.id === currentEditId)?.status || '')).toLowerCase() === 'rejected_by_rsa')
        ? '✅ Documents re-uploaded and returned to RSA.'
        : '✅ Documents re-uploaded and sent back for reviewer action.',
      'success'
    );
    closeModal(editModal);
  } catch (error) {
    showNotification('Update failed: ' + error.message, 'error');
  } finally {
    submitEditBtn.innerHTML = '<i class="fas fa-upload"></i> Re-upload Documents';
    updateSubmitButton();
  }
}

window.openDraftSubmission = async (id) => {
  const sub = allSubmissions.find((item) => item.id === id && String(item.status || '').toLowerCase() === 'draft');
  if (!sub) {
    showNotification('Draft not found.', 'error');
    return;
  }

  currentEditId = null;
  currentDraftId = id;
  currentHouseNumberReserved = true;
  currentDocType = null;
  currentFile = null;
  if (singleFileInput) singleFileInput.value = '';
  clearSingleFilePreview();
  await loadApprovedAgents();
  syncCurrentSubmissionAgentFallbackFromSubmission(sub);
  populateApprovedAgentSelect();
  applyDraftFormValues(sub);
  hydrateEditableSubmissionRequiredFields();
  hydrateCurrentUploads(sub.documents || []);
  if (customerAgentSelect) customerAgentSelect.value = getStoredAgentSelectionValue(sub);
  enableDraftEditingState();
  setUploadModalHeading('draft');
  syncUploadModalControlState('draft');
  if (saveDetailsBtn) {
    saveDetailsBtn.innerHTML = '<i class="fas fa-check"></i> Details Saved';
  }
  if (window.renderDocumentGridUpload) {
    window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'edit');
    window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'edit');
  }
  updateSubmitButton();
  uploadModal.classList.add('active');
};

async function downloadBulkImportTemplate() {
  if (!window.ExcelJS) {
    showNotification('Excel template library is not available right now.', 'error');
    return;
  }

  try {
    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Submission Drafts');
    const systemSettings = await getSystemSettings(db, { force: true });
    const headers = getCustomerBulkImportHeaders(systemSettings);
    sheet.addRow(headers);
    sheet.addRow([
      'John Doe', '1986-05-18', '0123456789', 'john@example.com', '08012345678', 'Agent Example',
      '12345678901', '12 Example Street, Abuja', 'Example Employer', String(currentUserProfile?.location || ''),
      '2026-04-27', 'Premium Pension', 'PEN123456789012', '2026-04-26', '12000000'
    ]);
    sheet.getRow(1).font = { bold: true };
    sheet.columns.forEach((column) => {
      column.width = 22;
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'cmbank_bulk_submission_template.xlsx';
    link.click();
    URL.revokeObjectURL(url);
    showNotification('Bulk template downloaded.', 'success');
  } catch (error) {
    showNotification('Could not generate bulk template.', 'error');
  }
}

function getCellText(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'object') {
    if (value.text) return String(value.text).trim();
    if (value.result !== undefined && value.result !== null) return String(value.result).trim();
  }
  return String(value).trim();
}

function normalizeImportHeader(header) {
  return String(header || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function getCustomerBulkImportHeaders(systemSettings = {}) {
  const configuredHeaders = Array.isArray(systemSettings.bulkImportRules?.requiredColumns) && systemSettings.bulkImportRules.requiredColumns.length
    ? systemSettings.bulkImportRules.requiredColumns
    : getDefaultSystemSettings().bulkImportRules.requiredColumns;
  return configuredHeaders.filter((header) => normalizeImportHeader(header) !== 'accountbank');
}

async function buildImportedDraftPayload(rowMap) {
  const rsaBalanceValue = getCellText(rowMap.rsabalance);
  const rsaBalance = parseMoney(rsaBalanceValue);
  const propertyRule = rsaBalance ? determinePropertyByRsa(rsaBalance) : null;
  let customerDetails = {
    name: getCellText(rowMap.customername),
    dob: getCellText(rowMap.dateofbirth),
    email: getCellText(rowMap.email),
    phone: getCellText(rowMap.phone),
    nin: getCellText(rowMap.nin),
    address: getCellText(rowMap.address),
    accountNo: getCellText(rowMap.accountno),
    accountBank: DEFAULT_CUSTOMER_ACCOUNT_BANK_NAME,
    accountBankCode: DEFAULT_CUSTOMER_ACCOUNT_BANK_CODE,
    employer: getCellText(rowMap.employer),
    originatingTP: getCellText(rowMap.originatingtp) || String(currentUserProfile?.location || '').trim(),
    mortgageLoanApplicationFormDate: getCellText(rowMap.mortgageformdate),
    pfa: getCellText(rowMap.pfa),
    penNo: getCellText(rowMap.penno),
    penNoNormalized: normalizePenNumber(getCellText(rowMap.penno)),
    rsaStatementDate: getCellText(rowMap.rsastatementdate),
    rsaBalance: rsaBalanceValue,
    rsa25: rsaBalance ? String(calculateRoundedRsa25(rsaBalance)) : '',
    propertyType: propertyRule?.name || '',
    houseNumber: '',
    tenor: '',
    propertyValue: propertyRule ? String(propertyRule.value) : '',
    facilityFee: propertyRule ? String(propertyRule.fee) : '',
    loanAmount: ''
  };
  customerDetails = await ensureStoredHouseNumber(customerDetails, { skipDomUpdate: true });

  return {
    customerName: customerDetails.name || 'Untitled Draft',
    customerDetails,
    uploadedBy: normalizeEmail(currentUser?.email),
    uploadedAt: serverTimestamp(),
    status: 'draft',
    comment: '',
    documents: [],
    documentTypes: [],
    houseNumber: customerDetails.houseNumber || '',
    penNoNormalized: normalizePenNumber(customerDetails.penNo),
    agentId: '',
    agentName: '',
    agentContactNumber: '',
    agentAccountNumber: '',
    agentAccountBank: '',
    importedAgentName: getCellText(rowMap.agentname),
    draftSource: 'excel',
    draftSavedAt: serverTimestamp()
  };
}

function findBulkImportDuplicateReason() {
  return '';
}

async function handleBulkImportSelection(event) {
  const file = event?.target?.files?.[0];
  if (!file) return;
  if (!window.ExcelJS) {
    showNotification('Excel import library is not available right now.', 'error');
    event.target.value = '';
    return;
  }

  try {
    window.showLoader?.('Reading Excel draft import...');
    const systemSettings = await getSystemSettings(db, { force: true });
    const requiredHeaders = getCustomerBulkImportHeaders(systemSettings);
    const workbook = new window.ExcelJS.Workbook();
    const buffer = await file.arrayBuffer();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error('No worksheet found in this Excel file.');

    const headerRow = sheet.getRow(1);
    const headerMap = new Map();
    headerRow.eachCell((cell, colNumber) => {
      const normalized = normalizeImportHeader(getCellText(cell.value));
      if (normalized) headerMap.set(colNumber, normalized);
    });
    const normalizedHeaders = new Set(Array.from(headerMap.values()));
    const missingHeaders = requiredHeaders
      .map((header) => normalizeImportHeader(header))
      .filter((header) => !normalizedHeaders.has(header));
    if (missingHeaders.length) {
      throw new Error(`Missing required bulk import columns: ${missingHeaders.join(', ')}`);
    }

    let importedCount = 0;
    const skipped = [];
    const importedPenNumbers = new Set();
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const rowMap = {};
      let hasAnyValue = false;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const key = headerMap.get(colNumber);
        if (!key) return;
        const value = getCellText(cell.value);
        if (value) hasAnyValue = true;
        rowMap[key] = value;
      });

      if (!hasAnyValue) continue;
      if (!getCellText(rowMap.customername)) {
        skipped.push('Missing Customer Name');
        continue;
      }
      const importedPenNo = normalizePenNumber(getCellText(rowMap.penno));
      if (importedPenNo && importedPenNumbers.has(importedPenNo)) {
        skipped.push(`Row ${rowNumber}: duplicate PEN number inside import file`);
        continue;
      }
      if (importedPenNo && !(await validatePenNumberAvailable(importedPenNo, ''))) {
        skipped.push(`Row ${rowNumber}: PEN number already exists`);
        continue;
      }
      const draftPayload = await buildImportedDraftPayload(rowMap);
      await addDoc(collection(db, 'submissions'), draftPayload);
      if (importedPenNo) importedPenNumbers.add(importedPenNo);
      importedCount += 1;
    }

    if (!importedCount && skipped.length) {
      showNotification(skipped[0], 'error');
    } else if (skipped.length) {
      showNotification(`Imported ${importedCount} draft(s). ${skipped.length} row(s) skipped.`, 'warning');
    } else {
      showNotification(`Imported ${importedCount} draft(s) successfully.`, 'success');
    }
  } catch (error) {
    showNotification('Bulk import failed: ' + error.message, 'error');
  } finally {
    window.hideLoader?.();
    if (event?.target) event.target.value = '';
  }
}

// ==================== LOAD SUBMISSIONS ====================
async function loadSubmissions() {
  const uploaderEmail = normalizeEmail(currentUser?.email);
  if (!uploaderEmail) return;

  submissionListenerUnsubs.forEach((unsubscribe) => {
    try { unsubscribe(); } catch (_) {}
  });
  submissionListenerUnsubs = [];
  submissionSnapshotSources.clear();

  const refreshMergedSubmissions = async () => {
    const merged = new Map();
    submissionSnapshotSources.forEach((rows) => {
      rows.forEach((row) => merged.set(row.id, row));
    });
    allSubmissions = Array.from(merged.values()).sort((a, b) => getSubmissionSortMillis(b) - getSubmissionSortMillis(a));
    const emails = new Set();
    allSubmissions.forEach((data) => {
      if (data.uploadedBy) emails.add(data.uploadedBy);
      if (data.reviewedBy) emails.add(data.reviewedBy);
      if (data.assignedTo) emails.add(data.assignedTo);
      if (data.assignedToRSA) emails.add(data.assignedToRSA);
      if (data.assignedToPayment) emails.add(data.assignedToPayment);
      if (data.paidBy) emails.add(data.paidBy);
      if (data.clearedBy) emails.add(data.clearedBy);
    });
    try { await ensureUserFullNames(Array.from(emails)); } catch (e) { }
    await renderRecentTable();
    renderDraftTable();
    await renderPendingTable();
    await renderApprovedTable();
    await renderRejectedTable();
    renderUploaderApplicationsTable();
    renderPaidTable();
    updateDashboardCards();
  };

  const attachSubmissionListener = (sourceKey, submissionQuery) => {
    const unsubscribe = onSnapshot(submissionQuery, async (snapshot) => {
      const rows = [];
      snapshot.forEach((docSnap) => {
        rows.push(normalizeSubmissionAgentFields({ id: docSnap.id, ...(docSnap.data() || {}) }));
      });
      submissionSnapshotSources.set(sourceKey, rows);
      await refreshMergedSubmissions();
    }, () => {
      showNotification('Error loading submissions', 'error');
    });
    submissionListenerUnsubs.push(unsubscribe);
  };

  attachSubmissionListener(
    'own',
    query(collection(db, 'submissions'), where('uploadedBy', '==', uploaderEmail), orderBy('uploadedAt', 'desc'))
  );

  registeredAgents
    .map((agent) => String(agent.id || '').trim())
    .filter(Boolean)
    .forEach((agentId) => {
      attachSubmissionListener(
        `agent-paid:${agentId}`,
        query(collection(db, 'submissions'), where('agentId', '==', agentId), where('status', '==', 'paid'))
      );
    });
}

async function ensureUserFullNames(emails) {
  if (!emails || emails.length === 0) return;
  await ensureUserFullNamesShared(db, emails);
  await Promise.all(emails.map(async (email) => {
    const normalized = normalizeEmail(email);
    if (!normalized || userFullNames.has(normalized)) return;
    const fullName = await getUserFullNameShared(db, normalized);
    userFullNames.set(normalized, fullName);
  }));
}

async function getUserProfileByEmail(email) {
  return getUserProfileByEmailShared(db, email);
}

function getSubmissionStageKey(submission) {
  const st = String(submission?.status || '').toLowerCase();
  if (st === 'cleared') return 'closed';
  if (st === 'sent_to_pfa' || st === 'rsa_submitted' || st === 'paid') return 'payment';
  if (st === 'processing_to_pfa' || st === 'approved') return 'rsa';
  return 'review';
}

function getCurrentHandlerEmail(submission) {
  const stage = getSubmissionStageKey(submission);
  if (stage === 'review') return normalizeEmail(submission.assignedTo || submission.reviewedBy);
  if (stage === 'rsa') return normalizeEmail(submission.assignedToRSA);
  if (stage === 'payment') return normalizeEmail(submission.assignedToPayment);
  return '';
}

async function renderStageContactLink(submission) {
  const handlerEmail = getCurrentHandlerEmail(submission);
  if (!handlerEmail) return '-';
  const profile = await getUserProfileByEmail(handlerEmail);
  const raw = String(profile?.whatsappNumber || profile?.phone || '').trim();
  return renderWhatsAppLink(raw);
}

async function getCurrentHandlerName(submission) {
  const handlerEmail = getCurrentHandlerEmail(submission);
  return handlerEmail ? await getUserFullName(handlerEmail) : 'Not assigned';
}

function getSubmissionSortMillis(submission) {
  return getStageTimestampMillis(getSubmissionCurrentStageEntryAt(submission));
}

function safeFormatDate(dateValue) {
  return formatAppDateTime(dateValue, 'N/A');
}

function normalizeWhatsAppPhone(raw) {
  const digits = String(raw || '').replace(/\D/g, '');
  if (!digits) return '';
  if (digits.startsWith('0') && digits.length === 11) return `234${digits.slice(1)}`;
  if (digits.length === 10) return `234${digits}`;
  if (digits.startsWith('234')) return digits;
  return digits;
}

function renderWhatsAppLink(raw) {
  const display = String(raw || '').trim();
  const normalized = normalizeWhatsAppPhone(display);
  if (!normalized) return '-';
  return `<a href="https://wa.me/${normalized}" target="_blank" rel="noopener noreferrer" aria-label="Open WhatsApp chat" title="Open WhatsApp chat"><i class="fab fa-whatsapp"></i></a>`;
}

async function renderRecentTable() {
  if (!recentTableBody) return;
  const recentItems = [...allSubmissions]
    .filter(isOwnUploaderSubmission)
    .filter((sub) => String(sub.status || '').toLowerCase() !== 'draft')
    .sort((a, b) => getSubmissionSortMillis(b) - getSubmissionSortMillis(a))
    .slice(0, 6);

  if (!recentItems.length) {
    recentTableBody.innerHTML = '<tr><td colspan="5" class="no-data">No recent applications</td></tr>';
    return;
  }

  recentTableBody.innerHTML = recentItems.map((sub) => `
    <tr>
      <td><strong>${escapeHtml(sub.customerName || '-')}</strong></td>
      <td>${safeFormatDate(getSubmissionCurrentStageEntryAt(sub))}</td>
      <td><span class="status-badge status-${String(sub.status || 'pending').toLowerCase().replace(/[^a-z0-9_-]+/g, '-')}">${escapeHtml(String(sub.status || 'pending').replace(/_/g, ' '))}</span></td>
      <td>${escapeHtml(sub.uploadedBy || '-')}</td>
      <td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> View</button> <button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td>
    </tr>
  `).join('');
}

function renderDraftTable() {
  if (!draftTableBody) return;
  const drafts = [...allSubmissions]
    .filter(isOwnUploaderSubmission)
    .filter((sub) => String(sub.status || '').toLowerCase() === 'draft')
    .sort((a, b) => getSubmissionSortMillis(b) - getSubmissionSortMillis(a));

  if (!drafts.length) {
    draftTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No saved drafts</td></tr>';
    return;
  }

  draftTableBody.innerHTML = drafts.map((sub) => {
    const docCount = Array.isArray(sub.documents) ? sub.documents.length : 0;
    const agentName = getSubmissionAgentDisplayName(sub);
    return `
      <tr data-submission-id="${sub.id}">
        <td><strong>${escapeHtml(sub.customerName || 'Untitled Draft')}</strong></td>
        <td>${escapeHtml(agentName)}</td>
        <td>${docCount}</td>
        <td>${safeFormatDate(sub.draftSavedAt || sub.uploadedAt)}</td>
        <td><span class="status-badge status-pending">Draft</span></td>
        <td>
          <button class="action-btn edit-btn" onclick="window.openDraftSubmission('${sub.id}')"><i class="fas fa-pen"></i> Resume</button>
          <button class="action-btn" onclick="window.deleteDraftSubmission('${sub.id}')" style="background:#b91c1c;color:#fff;border:none;"><i class="fas fa-trash"></i> Delete</button>
        </td>
      </tr>
    `;
  }).join('');
}

window.deleteDraftSubmission = async (id) => {
  const draft = allSubmissions.find((item) => item.id === id && String(item.status || '').toLowerCase() === 'draft');
  if (!draft) {
    showNotification('Draft not found.', 'error');
    return;
  }

  const draftName = String(draft.customerName || 'Untitled Draft').trim();
  const shouldDelete = window.confirm(`Delete draft "${draftName}"? This cannot be undone.`);
  if (!shouldDelete) return;

  try {
    await deleteDoc(doc(db, 'submissions', id));
    if (currentDraftId === id) {
      currentDraftId = null;
    }
    if (currentEditId === id) {
      currentEditId = null;
    }
    showNotification('Draft deleted successfully.', 'success');
  } catch (error) {
    showNotification('Failed to delete draft.', 'error');
  }
};

function updateDashboardCards() {
  const ownSubmissions = allSubmissions.filter(isOwnUploaderSubmission);
  const draft = ownSubmissions.filter(s => String(s.status || '').toLowerCase() === 'draft').length;
  const pending = ownSubmissions.filter(s => s.status === 'pending').length;
  const approved = ownSubmissions.filter((s) => getUploaderApplicationBucket(s) === 'approved').length;
  const applicationCounts = getUploaderApplicationCounts();
  const rejected = applicationCounts.rejected;
  const paid = applicationCounts.paid;
  const applicationsTotal = Object.values(applicationCounts).reduce((sum, value) => sum + value, 0);
  document.getElementById('cardPendingCount').textContent = pending;
  document.getElementById('cardApprovedCount').textContent = approved;
  document.getElementById('cardRejectedCount').textContent = rejected;
  const setBadge = (id, count) => {
    const badge = document.getElementById(id);
    if (badge) {
      badge.textContent = String(count);
      badge.style.display = 'inline-block';
    }
  };
  setBadge('draftCount', draft);
  setBadge('pendingCount', pending);
  setBadge('approvedCount', approved);
  setBadge('rejectedCount', rejected);
  setBadge('applicationsCount', applicationsTotal);
  setBadge('paidCount', paid);
  renderUploaderApplicationBadges();
}

function isUploaderApprovedLifecycleStatus(submission = {}) {
  const status = String(submission.status || '').toLowerCase();
  return (
    status === 'processing_to_pfa' ||
    status === 'approved' ||
    status === 'sent_to_pfa' ||
    status === 'rsa_submitted' ||
    status === 'paid' ||
    status === 'cleared' ||
    submission.finalSubmitted === true ||
    submission.rsaSubmitted === true
  );
}

function isUploaderSentToPfaStatus(submission = {}) {
  const status = String(submission.status || '').toLowerCase();
  return (
    status === 'sent_to_pfa' ||
    status === 'rsa_submitted' ||
    (
      (submission.finalSubmitted === true || submission.rsaSubmitted === true) &&
      status !== 'paid' &&
      status !== 'cleared'
    )
  );
}

function getUploaderApplicationBucket(submission = {}) {
  const status = String(submission.status || '').toLowerCase();
  const auditStatus = String(submission.auditCommissionStatus || '').toLowerCase();
  if (status === 'draft') return 'draft';
  if (status === 'pending') return 'pending';
  if (auditStatus === 'rejected') return 'rejected';
  if (status === 'rejected' || status === 'rejected_by_rsa') return 'rejected';
  if (status === 'paid') return 'paid';
  if (status === 'cleared') return 'cleared';
  if (submission.paymentMadeByUploader === true && auditStatus === 'pending') return 'audit';
  if (isUploaderSentToPfaStatus(submission)) return 'sent_to_pfa';
  if (status === 'approved' || status === 'processing_to_pfa') return 'approved';
  return '';
}

function isAuditCommissionRejected(submission = {}) {
  return String(submission.auditCommissionStatus || '').toLowerCase() === 'rejected';
}

function isAuditApplicationFrozen(submission = {}) {
  return submission.auditFrozen === true;
}

function isOwnUploaderSubmission(submission = {}) {
  return normalizeEmail(submission.uploadedBy) === normalizeEmail(currentUser?.email);
}

function isLinkedAgentPaidSubmission(submission = {}) {
  if (String(submission.status || '').toLowerCase() !== 'paid') return false;
  const agentId = String(submission.agentId || '').trim();
  return !!agentId && registeredAgents.some((agent) => String(agent.id || '').trim() === agentId);
}

function isVisibleUploaderApplicationForTab(submission = {}, tab = currentUploaderApplicationTab) {
  if (tab === 'paid') return isOwnUploaderSubmission(submission) || isLinkedAgentPaidSubmission(submission);
  return isOwnUploaderSubmission(submission);
}

function getActiveCommissionScopeLabel(scope = currentUploaderPaidScope) {
  return scope === 'others' ? 'Other Users Applications' : 'My Applications';
}

function getUploaderApplicationCounts() {
  return allSubmissions.reduce((acc, sub) => {
    const bucket = getUploaderApplicationBucket(sub);
    if (bucket && Object.prototype.hasOwnProperty.call(acc, bucket) && isVisibleUploaderApplicationForTab(sub, bucket)) acc[bucket] += 1;
    return acc;
  }, { draft: 0, pending: 0, approved: 0, rejected: 0, sent_to_pfa: 0, audit: 0, paid: 0, cleared: 0 });
}

function getSubmissionPfaName(submission = {}) {
  return String(submission?.customerDetails?.pfa || submission?.pfa || '').trim() || '-';
}

function getUploaderAuditNote(submission = {}) {
  const auditStatus = String(submission.auditCommissionStatus || '').toLowerCase();
  if (auditStatus === 'pending') return 'Payment made submitted to Audit';
  if (auditStatus === 'accepted') return 'Accepted by Audit';
  if (auditStatus === 'rejected') {
    return `Rejected by Audit: ${submission.auditCommissionRejectionReason || 'No reason provided'}`;
  }
  return '-';
}

function getUploaderPaymentStageEntryAt(submission = {}) {
  const bucket = getUploaderApplicationBucket(submission);
  if (bucket === 'paid') return getSubmissionPaidEntryAt(submission);
  if (bucket === 'cleared') return getSubmissionClearedEntryAt(submission);
  if (bucket === 'audit') return submission.auditCommissionResubmittedAt || submission.auditCommissionSubmittedAt || submission.paymentMadeAt || getSubmissionPaymentEntryAt(submission);
  if (bucket === 'sent_to_pfa') return getSubmissionPaymentEntryAt(submission);
  return getSubmissionCurrentStageEntryAt(submission);
}

async function getUploaderPaymentResidentOfficer(submission = {}) {
  const bucket = getUploaderApplicationBucket(submission);
  const auditStatus = String(submission.auditCommissionStatus || '').toLowerCase();

  if (bucket === 'audit' || (bucket === 'sent_to_pfa' && auditStatus === 'pending')) return 'Audit';
  if (bucket === 'sent_to_pfa' && submission.paymentMadeByUploader !== true) return 'Uploader';

  const officerEmail =
    bucket === 'cleared'
      ? (submission.clearedBy || submission.assignedToPayment || submission.paidBy || '')
      : bucket === 'paid'
        ? (submission.paidBy || submission.assignedToPayment || '')
        : (submission.assignedToPayment || submission.paidBy || submission.finalSubmittedBy || submission.rsaSubmittedBy || '');

  const normalized = normalizeEmail(officerEmail);
  if (!normalized) return bucket === 'cleared' ? 'Closed' : '-';
  return getUserFullName(normalized);
}

function getUploaderApplicationRows(tab = currentUploaderApplicationTab) {
  const search = String(applicationsSearch?.value || '').trim().toLowerCase();
  return allSubmissions
    .filter((sub) => getUploaderApplicationBucket(sub) === tab)
    .filter((sub) => isVisibleUploaderApplicationForTab(sub, tab))
    .filter((sub) => {
      if (tab !== 'paid') return true;
      const isMine = isOwnUploaderSubmission(sub);
      return currentUploaderPaidScope === 'others' ? !isMine : isMine;
    })
    .filter((sub) => {
      if (!search) return true;
      return [
        sub.customerName,
        getSubmissionAgentDisplayName(sub),
        getSubmissionPfaName(sub),
        formatSubmissionStatusLabel(sub.status || ''),
        getUploaderAuditNote(sub)
      ].some((value) => String(value || '').toLowerCase().includes(search));
    })
    .sort((a, b) => getSubmissionSortMillis(b) - getSubmissionSortMillis(a));
}

function renderUploaderApplicationBadges() {
  const counts = getUploaderApplicationCounts();
  const badgeMap = {
    appDraftCount: counts.draft,
    appPendingCount: counts.pending,
    appApprovedCount: counts.approved,
    appRejectedCount: counts.rejected,
    appSentToPfaCount: counts.sent_to_pfa,
    appAuditCount: counts.audit,
    appPaidCount: counts.paid,
    appClearedCount: counts.cleared
  };
  Object.entries(badgeMap).forEach(([id, value]) => {
    const badge = document.getElementById(id);
    if (badge) badge.textContent = String(value);
  });
}

function setUploaderApplicationsColumns(columns = []) {
  if (!applicationsTableHeadRow) return;
  applicationsTableHeadRow.innerHTML = columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('');
}

function getUploaderApplicationColumnCount() {
  return applicationsTableHeadRow?.querySelectorAll('th')?.length || 7;
}

function renderUploaderApplicationsEmpty(label) {
  if (!applicationsTableBody) return;
  const text = currentUploaderApplicationTab === 'paid'
    ? `No ${label}`
    : `No ${label} applications`;
  applicationsTableBody.innerHTML = `<tr><td colspan="${getUploaderApplicationColumnCount()}" class="no-data">${escapeHtml(text)}</td></tr>`;
}

function getUploaderSubmissionDetailsButtonHtml(submissionId, label = 'View') {
  return `<button class="action-btn view-btn-small" onclick="window.viewSubmissionDetails('${submissionId}')"><i class="fas fa-circle-info"></i> ${escapeHtml(label)}</button>`;
}

function getUploaderSubmissionDocsButtonHtml(submissionId, label = 'Docs') {
  return `<button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${submissionId}')"><i class="fas fa-eye"></i> ${escapeHtml(label)}</button>`;
}

async function renderUploaderApplicationsTable() {
  renderUploaderApplicationBadges();
  if (!applicationsTableBody) return;

  const activeCommissionScopeStrip = document.getElementById('activeCommissionScopeStrip');
  if (activeCommissionScopeStrip) {
    activeCommissionScopeStrip.style.display = currentUploaderApplicationTab === 'paid' ? 'flex' : 'none';
    activeCommissionScopeStrip.querySelectorAll('[data-active-commission-scope]').forEach((button) => {
      button.classList.toggle('active', button.dataset.activeCommissionScope === currentUploaderPaidScope);
    });
  }

  const rows = getUploaderApplicationRows();

  if (currentUploaderApplicationTab === 'draft') {
    setUploaderApplicationsColumns(['Customer Name', 'Agent', 'Documents', 'Last Saved', 'Action']);
    if (!rows.length) {
      renderUploaderApplicationsEmpty('draft');
      return;
    }
    applicationsTableBody.innerHTML = rows.map((sub) => {
      const docCount = Array.isArray(sub.documents) ? sub.documents.length : 0;
      return `
        <tr data-submission-id="${sub.id}">
          <td><strong>${escapeHtml(sub.customerName || 'Untitled Draft')}</strong></td>
          <td>${escapeHtml(getSubmissionAgentDisplayName(sub))}</td>
          <td>${docCount}</td>
          <td>${safeFormatDate(sub.draftSavedAt || sub.uploadedAt)}</td>
          <td>
            ${getUploaderSubmissionDetailsButtonHtml(sub.id)}
            <button class="action-btn edit-btn" onclick="window.openDraftSubmission('${sub.id}')"><i class="fas fa-pen"></i> Resume</button>
            <button class="action-btn" onclick="window.deleteDraftSubmission('${sub.id}')" style="background:#b91c1c;color:#fff;border:none;"><i class="fas fa-trash"></i> Delete</button>
          </td>
        </tr>
      `;
    }).join('');
    return;
  }

  if (currentUploaderApplicationTab === 'pending') {
    setUploaderApplicationsColumns(['Customer Name', 'Agent', 'Contact (WhatsApp)', 'Current Handler', 'Current Stage Entry', 'Comment', 'View', 'Track']);
    if (!rows.length) {
      renderUploaderApplicationsEmpty('pending');
      return;
    }
    let html = '';
    for (const sub of rows) {
      const date = safeFormatDate(getSubmissionReviewEntryAt(sub));
      const assignedName = await getCurrentHandlerName(sub);
      const whatsapp = await renderStageContactLink(sub);
      const agentName = escapeHtml(getSubmissionAgentDisplayName(sub));
      html += `<tr data-submission-id="${sub.id}"><td><strong>${escapeHtml(sub.customerName || '-')}</strong></td><td>${agentName}</td><td>${whatsapp}</td><td>${escapeHtml(assignedName || '-')}</td><td>${date}</td><td>${escapeHtml(sub.comment || '-')}</td><td>${getUploaderSubmissionDetailsButtonHtml(sub.id)} ${getUploaderSubmissionDocsButtonHtml(sub.id)} <button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
    }
    applicationsTableBody.innerHTML = html;
    return;
  }

  if (currentUploaderApplicationTab === 'approved') {
    setUploaderApplicationsColumns(['Customer Name', 'Agent', 'Contact (WhatsApp)', 'Assigned To', 'Upload Date/Time', 'Approved By', 'Approved Date/Time', 'View', 'Track']);
    if (!rows.length) {
      renderUploaderApplicationsEmpty('approved');
      return;
    }
    let html = '';
    for (const sub of rows) {
      const uploadDate = safeFormatDate(sub.uploadedAt);
      const approvedDate = safeFormatDate(getSubmissionApprovalEntryAt(sub));
      const approvedByKey = normalizeEmail(sub.reviewedBy);
      const approvedBy = (approvedByKey && userFullNames.get(approvedByKey)) ? userFullNames.get(approvedByKey) : (sub.reviewedBy || '-');
      const assignedName = await getCurrentHandlerName(sub);
      const whatsapp = await renderStageContactLink(sub);
      const agentName = escapeHtml(getSubmissionAgentDisplayName(sub));
      html += `<tr data-submission-id="${sub.id}"><td><strong>${escapeHtml(sub.customerName || '-')}</strong></td><td>${agentName}</td><td>${whatsapp}</td><td>${escapeHtml(assignedName || '-')}</td><td>${uploadDate}</td><td>${escapeHtml(approvedBy || '-')}</td><td>${approvedDate}</td><td>${getUploaderSubmissionDetailsButtonHtml(sub.id)} ${getUploaderSubmissionDocsButtonHtml(sub.id)} <button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
    }
    applicationsTableBody.innerHTML = html;
    return;
  }

  if (currentUploaderApplicationTab === 'rejected') {
    setUploaderApplicationsColumns(['Customer Name', 'Agent', 'Chat', 'Re-upload Count', 'Assigned To', 'Upload Date/Time', 'Rejection Details', 'Action', 'View', 'Track']);
    if (!rows.length) {
      renderUploaderApplicationsEmpty('rejected');
      return;
    }
    let html = '';
    for (const sub of rows) {
      const auditRejected = isAuditCommissionRejected(sub);
      const auditFrozen = isAuditApplicationFrozen(sub);
      const fixCount = auditRejected ? Number(sub.auditCommissionResubmitCount || 0) : Number(sub.fixCount || 0);
      const date = safeFormatDate(auditRejected ? (sub.auditCommissionRejectedAt || sub.auditCommissionSubmittedAt || sub.paymentMadeAt) : getSubmissionReviewEntryAt(sub));
      const assignedName = auditRejected ? 'Audit' : (sub.assignedTo ? await getUserFullName(sub.assignedTo) : 'Not assigned');
      const agentName = escapeHtml(getSubmissionAgentDisplayName(sub));
      const chatBtn = auditFrozen
        ? `<button class="action-btn app-chat-trigger" disabled title="Application frozen by Audit"><i class="fas fa-lock"></i> Frozen</button>`
        : `<button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button>`;
      const reasonBtn = hasRejectionHistory(sub)
        ? `<button class="action-btn reason-btn" onclick="window.openUploaderRejectionReasonModal('${sub.id}')"><i class="fas fa-eye"></i> View Details</button>`
        : 'No reason provided';
      const actionCell = auditFrozen
        ? `<div class="audit-frozen-uploader-notice"><i class="fas fa-snowflake"></i><span>Frozen by Audit</span><small>Wait for Audit to unfreeze this application.</small></div>`
        : auditRejected
        ? `<div class="audit-rejection-actions">
            <button class="action-btn edit-btn" onclick="window.openAuditPaymentResubmitModal('${sub.id}')" title="Submit correction to Audit"><i class="fas fa-paper-plane"></i> Resubmit</button>
            <button class="action-btn dissolve-btn" onclick="window.dissolveAuditPaymentRequest('${sub.id}')" title="Return to Sent to PFA"><i class="fas fa-rotate-left"></i> Dissolve</button>
          </div>`
        : `<button class="action-btn edit-btn" onclick="window.openEditModal('${sub.id}')" title="Correction count: ${fixCount}"><i class="fas fa-edit"></i> Re-upload</button>`;
      html += `<tr data-submission-id="${sub.id}"><td><strong>${escapeHtml(sub.customerName || '-')}</strong></td><td>${agentName}</td><td>${chatBtn}</td><td>${fixCount}</td><td>${escapeHtml(assignedName || '-')}</td><td>${date}</td><td>${reasonBtn}</td><td>${actionCell}</td><td>${getUploaderSubmissionDetailsButtonHtml(sub.id)} ${getUploaderSubmissionDocsButtonHtml(sub.id)}</td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
    }
    applicationsTableBody.innerHTML = html;
    return;
  }

  setUploaderApplicationsColumns(
    currentUploaderApplicationTab === 'paid'
      ? ['Customer Name', 'Agent', 'PFA', 'Uploaded By', 'Time Entered', 'Current Officer', 'Action']
      : ['Customer Name', 'Agent', 'PFA', 'Time Entered', 'Current Officer', 'Action']
  );
  if (!rows.length) {
    const label = currentUploaderApplicationTab === 'paid'
      ? getActiveCommissionScopeLabel()
      : currentUploaderApplicationTab.replace(/_/g, ' ');
    renderUploaderApplicationsEmpty(label);
    return;
  }

  let html = '';
  for (const sub of rows) {
    const status = String(sub.status || '').toLowerCase();
    const auditStatus = String(sub.auditCommissionStatus || '').toLowerCase();
    const canReportPayment = currentUploaderApplicationTab === 'sent_to_pfa' && auditStatus !== 'pending';
    const paymentButton = currentUploaderApplicationTab === 'sent_to_pfa'
      ? `<button class="action-btn view-btn-small" ${canReportPayment ? '' : 'disabled'} onclick="window.markUploaderPaymentMade('${sub.id}')"><i class="fas fa-money-bill-wave"></i> ${auditStatus === 'pending' ? 'Reported' : 'Payment Made'}</button>`
      : '';
    const auditLabel = currentUploaderApplicationTab === 'audit'
      ? '<span class="audit-pending-pill"><i class="fas fa-hourglass-half"></i> Pending Audit</span>'
      : '';
    const residentOfficer = await getUploaderPaymentResidentOfficer(sub);
    const uploadedByCell = currentUploaderApplicationTab === 'paid'
      ? `<td>${escapeHtml(isOwnUploaderSubmission(sub) ? 'Me' : (sub.uploadedBy || '-'))}</td>`
      : '';
    html += `
      <tr data-submission-id="${sub.id}">
        <td><strong>${escapeHtml(sub.customerName || '-')}</strong></td>
        <td>${escapeHtml(getSubmissionAgentDisplayName(sub))}</td>
        <td>${escapeHtml(getSubmissionPfaName(sub))}</td>
        ${uploadedByCell}
        <td>${safeFormatDate(getUploaderPaymentStageEntryAt(sub))}</td>
        <td>${escapeHtml(residentOfficer)}</td>
        <td>
          ${paymentButton}
          ${auditLabel}
          ${getUploaderSubmissionDetailsButtonHtml(sub.id)}
          ${getUploaderSubmissionDocsButtonHtml(sub.id)}
          <button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button>
        </td>
      </tr>
    `;
  }
  applicationsTableBody.innerHTML = html;
}

function switchUploaderApplicationTab(tab = 'pending') {
  currentUploaderApplicationTab = UPLOADER_APPLICATION_TABS.includes(tab) ? tab : 'draft';
  rememberUploaderTab('applications', currentUploaderApplicationTab);
  if (currentUploaderApplicationTab !== 'paid') currentUploaderPaidScope = 'mine';
  document.querySelectorAll('[data-uploader-application-tab]').forEach((button) => {
    button.classList.toggle('active', button.dataset.uploaderApplicationTab === currentUploaderApplicationTab);
  });
  renderUploaderApplicationsTable();
}

function showPaymentMadeConfirmation(submission = {}) {
  return new Promise((resolve) => {
    const pfaName = getSubmissionPfaName(submission);
    const customerName = String(submission.customerName || 'this customer').trim();
    const modal = document.createElement('div');
    modal.className = 'modal active payment-made-confirm-modal';
    modal.innerHTML = `
      <div class="modal-content payment-made-confirm-card">
        <div class="payment-made-confirm-icon">
          <i class="fas fa-money-check-dollar"></i>
        </div>
        <h2>Confirm Payment Made</h2>
        <p>Are you sure payment has been made by <strong>${escapeHtml(pfaName)}</strong> for <strong>${escapeHtml(customerName)}</strong>?</p>
        <div class="payment-made-confirm-actions">
          <button type="button" class="cancel-btn" data-confirm-payment="no">Cancel</button>
          <button type="button" class="submit-btn" data-confirm-payment="yes">
            <i class="fas fa-check"></i> Confirm
          </button>
        </div>
      </div>
    `;
    const close = (value) => {
      modal.remove();
      resolve(value);
    };
    modal.addEventListener('click', (event) => {
      if (event.target === modal) close(false);
      const button = event.target.closest('[data-confirm-payment]');
      if (!button) return;
      close(button.dataset.confirmPayment === 'yes');
    });
    document.body.appendChild(modal);
  });
}

function showAuditPaymentResubmitDialog(submission = {}) {
  return new Promise((resolve) => {
    const modal = document.createElement('div');
    modal.className = 'modal active audit-payment-resubmit-modal';
    modal.innerHTML = `
      <div class="modal-content audit-payment-resubmit-card">
        <div class="payment-made-confirm-icon">
          <i class="fas fa-file-circle-plus"></i>
        </div>
        <h2>Resubmit Payment Request</h2>
        <p>Add a correction comment and upload the supporting document for <strong>${escapeHtml(submission.customerName || 'this application')}</strong>.</p>
        <label class="audit-resubmit-field">
          <span>Correction Comment</span>
          <textarea id="auditPaymentResubmitComment" rows="4" placeholder="Enter correction comment"></textarea>
        </label>
        <label class="audit-resubmit-field">
          <span>Correction Document</span>
          <input id="auditPaymentResubmitFile" type="file" accept=".pdf,image/*">
        </label>
        <div class="payment-made-confirm-actions">
          <button type="button" class="cancel-btn" data-audit-resubmit="cancel">Cancel</button>
          <button type="button" class="submit-btn" data-audit-resubmit="submit">
            <i class="fas fa-paper-plane"></i> Resubmit
          </button>
        </div>
      </div>
    `;
    const close = (value) => {
      modal.remove();
      resolve(value);
    };
    modal.addEventListener('click', (event) => {
      if (event.target === modal) {
        close(null);
        return;
      }
      const button = event.target.closest('[data-audit-resubmit]');
      if (!button) return;
      if (button.dataset.auditResubmit === 'cancel') {
        close(null);
        return;
      }
      const commentEl = modal.querySelector('#auditPaymentResubmitComment');
      const fileEl = modal.querySelector('#auditPaymentResubmitFile');
      const comment = String(commentEl?.value || '').trim();
      const file = fileEl?.files?.[0] || null;
      commentEl?.classList.toggle('invalid', !comment);
      fileEl?.classList.toggle('invalid', !file);
      if (!comment || !file) return;
      close({ comment, file });
    });
    document.body.appendChild(modal);
    setTimeout(() => modal.querySelector('#auditPaymentResubmitComment')?.focus(), 0);
  });
}

function showAuditDissolveConfirmation(submission = {}) {
  return new Promise((resolve) => {
    const modal = document.createElement('div');
    modal.className = 'modal active payment-made-confirm-modal';
    modal.innerHTML = `
      <div class="modal-content payment-made-confirm-card dissolve-confirm-card">
        <div class="payment-made-confirm-icon dissolve">
          <i class="fas fa-rotate-left"></i>
        </div>
        <h2>Dissolve Payment Request</h2>
        <p>This will return <strong>${escapeHtml(submission.customerName || 'this application')}</strong> back to Sent to PFA.</p>
        <div class="payment-made-confirm-actions">
          <button type="button" class="cancel-btn" data-audit-dissolve="no">Cancel</button>
          <button type="button" class="submit-btn danger" data-audit-dissolve="yes">
            <i class="fas fa-check"></i> Dissolve
          </button>
        </div>
      </div>
    `;
    const close = (value) => {
      modal.remove();
      resolve(value);
    };
    modal.addEventListener('click', (event) => {
      if (event.target === modal) close(false);
      const button = event.target.closest('[data-audit-dissolve]');
      if (!button) return;
      close(button.dataset.auditDissolve === 'yes');
    });
    document.body.appendChild(modal);
  });
}

window.markUploaderPaymentMade = async (submissionId) => {
  if (!assertWritable('Payment made report')) return;
  const sub = allSubmissions.find((item) => item.id === submissionId);
  if (!sub) {
    showNotification('Application not found', 'error');
    return;
  }
  if (!isUploaderSentToPfaStatus(sub)) {
    showNotification('Only Sent to PFA applications can be reported as payment made.', 'warning');
    return;
  }

  const pfaName = getSubmissionPfaName(sub);
  const confirmed = await showPaymentMadeConfirmation(sub);
  if (!confirmed) return;

  try {
    await updateDoc(doc(db, 'submissions', submissionId), {
      paymentMadeByUploader: true,
      paymentMadeAt: serverTimestamp(),
      paymentMadeBy: currentUser?.email || '',
      auditCommissionStatus: 'pending',
      auditCommissionRejectionReason: '',
      auditCommissionSubmittedAt: serverTimestamp(),
      auditCommissionSubmittedBy: currentUser?.email || '',
      updatedAt: serverTimestamp()
    });

    await addDoc(collection(db, 'audit'), {
      action: 'uploader_payment_made_reported',
      submissionId,
      customerName: sub.customerName || '',
      pfaName,
      performedBy: currentUser?.email || '',
      timestamp: serverTimestamp()
    }).catch(() => {});

    await notifyAdminPushEvent({
      currentUser,
      eventType: 'commission_payment_reported',
      title: 'Payment Request Submitted',
      body: `${sub.customerName || 'An application'} was submitted for payment approval.`,
      clickUrl: '/reports-monitoring-dashboard.html',
      meta: {
        submissionId,
        customerName: sub.customerName || '',
        pfaName,
        uploadedBy: sub.uploadedBy || '',
        submittedBy: currentUser?.email || ''
      }
    }).catch(() => {});

    showNotification('Payment made report sent to Audit.', 'success');
    switchUploaderApplicationTab('audit');
  } catch (error) {
    showNotification('Failed to report payment made.', 'error');
  }
};

window.openAuditPaymentResubmitModal = async (submissionId) => {
  if (!assertWritable('Payment request resubmit')) return;
  const sub = allSubmissions.find((item) => item.id === submissionId);
  if (!sub) {
    showNotification('Application not found', 'error');
    return;
  }
  if (isAuditApplicationFrozen(sub)) {
    showNotification('This application is frozen by Audit. It must be unfrozen before you can resubmit it.', 'warning');
    return;
  }
  if (!isAuditCommissionRejected(sub)) {
    showNotification('Only rejected Audit payment requests can be resubmitted.', 'warning');
    return;
  }

  const result = await showAuditPaymentResubmitDialog(sub);
  if (!result) return;

  try {
    const storage = new BackblazeStorage();
    showNotification('Uploading correction document...', 'info');
    const uploadResult = await storage.uploadFile(result.file, sub.customerName || 'application', 'audit_payment_correction');
    const uploadedAt = await getTrustedNowIso();
    const correctionDocument = {
      name: result.file.name,
      fileName: uploadResult.fileName || result.file.name,
      fileId: uploadResult.fileId || '',
      fileUrl: uploadResult.fileUrl || '',
      documentType: 'audit_payment_correction',
      comment: result.comment,
      uploadedAt,
      uploadedBy: currentUser?.email || ''
    };
    const nextCount = Number(sub.auditCommissionResubmitCount || 0) + 1;

    await updateDoc(doc(db, 'submissions', submissionId), {
      paymentMadeByUploader: true,
      paymentMadeAt: serverTimestamp(),
      paymentMadeBy: currentUser?.email || '',
      auditCommissionStatus: 'pending',
      auditCommissionRejectionReason: '',
      auditCommissionSubmittedAt: serverTimestamp(),
      auditCommissionSubmittedBy: currentUser?.email || '',
      auditCommissionResubmittedAt: serverTimestamp(),
      auditCommissionResubmittedBy: currentUser?.email || '',
      auditCommissionResubmitComment: result.comment,
      auditCommissionResubmitCount: nextCount,
      auditCommissionCorrectionDocuments: arrayUnion(correctionDocument),
      updatedAt: serverTimestamp()
    });

    await addDoc(collection(db, 'audit'), {
      action: 'uploader_payment_request_resubmitted',
      submissionId,
      customerName: sub.customerName || '',
      comment: result.comment,
      correctionDocumentName: result.file.name,
      performedBy: currentUser?.email || '',
      timestamp: serverTimestamp()
    }).catch(() => {});

    await notifyAdminPushEvent({
      currentUser,
      eventType: 'commission_payment_reported',
      title: 'Payment Request Resubmitted',
      body: `${sub.customerName || 'An application'} was resubmitted for Audit payment approval.`,
      clickUrl: '/reports-monitoring-dashboard.html',
      meta: {
        submissionId,
        customerName: sub.customerName || '',
        uploadedBy: sub.uploadedBy || '',
        submittedBy: currentUser?.email || '',
        correctionDocumentName: result.file.name
      }
    }).catch(() => {});

    showNotification('Payment request resubmitted to Audit.', 'success');
    switchUploaderApplicationTab('audit');
  } catch (error) {
    showNotification(`Failed to resubmit payment request: ${error.message || error}`, 'error');
  }
};

window.dissolveAuditPaymentRequest = async (submissionId) => {
  if (!assertWritable('Dissolve payment request')) return;
  const sub = allSubmissions.find((item) => item.id === submissionId);
  if (!sub) {
    showNotification('Application not found', 'error');
    return;
  }
  if (isAuditApplicationFrozen(sub)) {
    showNotification('This application is frozen by Audit. It must be unfrozen before you can dissolve it.', 'warning');
    return;
  }
  if (!isAuditCommissionRejected(sub)) {
    showNotification('Only rejected Audit payment requests can be dissolved.', 'warning');
    return;
  }

  const confirmed = await showAuditDissolveConfirmation(sub);
  if (!confirmed) return;

  try {
    await updateDoc(doc(db, 'submissions', submissionId), {
      paymentMadeByUploader: false,
      auditCommissionStatus: 'dissolved',
      auditCommissionRejectionReason: '',
      auditCommissionDissolvedAt: serverTimestamp(),
      auditCommissionDissolvedBy: currentUser?.email || '',
      updatedAt: serverTimestamp()
    });

    await addDoc(collection(db, 'audit'), {
      action: 'uploader_payment_request_dissolved',
      submissionId,
      customerName: sub.customerName || '',
      performedBy: currentUser?.email || '',
      timestamp: serverTimestamp()
    }).catch(() => {});

    showNotification('Payment request returned to Sent to PFA.', 'success');
    switchUploaderApplicationTab('sent_to_pfa');
  } catch (error) {
    showNotification('Failed to dissolve payment request.', 'error');
  }
};

async function renderPendingTable() {
  if (!pendingTableBody) { return; }
  const pending = allSubmissions
    .filter(isOwnUploaderSubmission)
    .filter(s => s.status === 'pending')
    .slice()
    .sort((a, b) => getStageTimestampMillis(getSubmissionReviewEntryAt(b)) - getStageTimestampMillis(getSubmissionReviewEntryAt(a)));
  if (pending.length === 0) { pendingTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No pending documents</td></tr>'; return; }
  let html = '';
  for (const sub of pending) {
    const date = safeFormatDate(getSubmissionReviewEntryAt(sub));
    const assignedName = await getCurrentHandlerName(sub);
    const whatsapp = await renderStageContactLink(sub);
    const agentName = escapeHtml(getSubmissionAgentDisplayName(sub));
    html += `<tr><td><strong>${escapeHtml(sub.customerName || '-')}</strong></td><td>${agentName}</td><td>${whatsapp}</td><td>${escapeHtml(assignedName || '-')}</td><td>${date}</td><td>${escapeHtml(sub.comment || '-')}</td><td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> View</button> <button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
  }
  pendingTableBody.innerHTML = html;
}

async function renderApprovedTable() {
  if (!approvedTableBody) { return; }
  const approved = allSubmissions
    .filter(isOwnUploaderSubmission)
    .filter((s) => isUploaderApprovedLifecycleStatus(s))
    .slice()
    .sort((a, b) => getStageTimestampMillis(getSubmissionApprovalEntryAt(b)) - getStageTimestampMillis(getSubmissionApprovalEntryAt(a)));
  if (approved.length === 0) { approvedTableBody.innerHTML = '<tr><td colspan="9" class="no-data">No approved documents</td></tr>'; return; }
  let html = '';
  for (const sub of approved) {
    const uploadDate = safeFormatDate(sub.uploadedAt);
    const approvedDate = safeFormatDate(getSubmissionApprovalEntryAt(sub));
    const approvedByKey = normalizeEmail(sub.reviewedBy);
    const approvedBy = (approvedByKey && userFullNames.get(approvedByKey)) ? userFullNames.get(approvedByKey) : (sub.reviewedBy || '-');
    const assignedName = await getCurrentHandlerName(sub);
    const whatsapp = await renderStageContactLink(sub);
    const agentName = escapeHtml(getSubmissionAgentDisplayName(sub));
    html += `<tr><td><strong>${escapeHtml(sub.customerName || '-')}</strong></td><td>${agentName}</td><td>${whatsapp}</td><td>${escapeHtml(assignedName || '-')}</td><td>${uploadDate}</td><td>${escapeHtml(approvedBy || '-')}</td><td>${approvedDate}</td><td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDetails('${sub.id}')"><i class="fas fa-circle-info"></i> Details</button> <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> Docs</button> <button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
  }
  approvedTableBody.innerHTML = html;
}

async function renderRejectedTable() {
  if (!rejectedTableBody) { return; }
  const rejected = allSubmissions
    .filter(isOwnUploaderSubmission)
    .filter(s => ['rejected', 'rejected_by_rsa'].includes(String(s.status || '').toLowerCase()))
    .slice()
    .sort((a, b) => getStageTimestampMillis(getSubmissionRejectionEntryAt(b)) - getStageTimestampMillis(getSubmissionRejectionEntryAt(a)));
  if (rejected.length === 0) { rejectedTableBody.innerHTML = '<tr><td colspan="10" class="no-data">No rejected documents</td></tr>'; return; }
  let html = '';
  for (const sub of rejected) {
    const fixCount = Number(sub.fixCount || 0);
    const date = safeFormatDate(getSubmissionReviewEntryAt(sub));
    const assignedName = sub.assignedTo ? await getUserFullName(sub.assignedTo) : 'Not assigned';
    const agentName = escapeHtml(getSubmissionAgentDisplayName(sub));
    const chatBtn = `<button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button>`;
    const reasonBtn = hasRejectionHistory(sub)
      ? `<button class="action-btn reason-btn" onclick="window.openUploaderRejectionReasonModal('${sub.id}')"><i class="fas fa-eye"></i> View Details</button>`
      : 'No reason provided';
    html += `<tr data-submission-id="${sub.id}"><td><strong>${escapeHtml(sub.customerName || '-')}</strong></td><td>${agentName}</td><td>${chatBtn}</td><td>${fixCount}</td><td>${escapeHtml(assignedName || '-')}</td><td>${date}</td><td>${reasonBtn}</td><td><button class="action-btn edit-btn" onclick="window.openEditModal('${sub.id}')" title="Correction count: ${fixCount}"><i class="fas fa-edit"></i> Re-upload</button></td><td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> View</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
  }
  rejectedTableBody.innerHTML = html;
}

window.openUploaderRejectionReasonModal = (submissionId) => {
  const sub = allSubmissions.find((s) => s.id === submissionId);
  if (!sub || !uploaderRejectionReasonModal) return;

  const entries = getRejectionHistoryEntries(sub);
  const contactValue = String(sub.customerDetails?.phone || sub.customerPhone || '-').trim() || '-';

  if (uploaderRejectionReasonCustomerName) {
    uploaderRejectionReasonCustomerName.textContent = sub.customerName || 'Unknown';
  }
  if (uploaderRejectionReasonContact) {
    uploaderRejectionReasonContact.textContent = `Contact: ${contactValue}`;
  }
  if (uploaderRejectionReasonHistory) {
    if (entries.length) {
      uploaderRejectionReasonHistory.innerHTML = `
        <ol class="rejection-history-list">
          ${entries.map((entry, index) => {
            const timeText = entry.rejectedAt ? safeFormatDate(entry.rejectedAt) : 'Time not available';
            const rejectedByKey = normalizeEmail(entry.rejectedBy || '');
            const rejectedBy = rejectedByKey && userFullNames.get(rejectedByKey)
              ? userFullNames.get(rejectedByKey)
              : (entry.rejectedBy || 'Not available');
            return `<li><strong>Rejection ${index + 1}</strong><div><b>Reason:</b> ${escapeHtml(entry.reason)}</div><div><b>Rejected By:</b> ${escapeHtml(rejectedBy)}</div><span class="rejection-history-time"><b>Rejected Date/Time:</b> ${escapeHtml(timeText)}</span></li>`;
          }).join('')}
        </ol>
      `;
    } else {
      uploaderRejectionReasonHistory.textContent = 'No rejection reason available.';
    }
    uploaderRejectionReasonHistory.style.display = 'block';
  }

  uploaderRejectionReasonModal.classList.add('active');
};

function getUploaderAgentCommissionUiRefs() {
  return {
    paidTab: document.getElementById('paidTab'),
    totalSentCard: document.getElementById('uploaderAgentCommissionTotalSent'),
    totalActiveCard: document.getElementById('uploaderAgentCommissionTotalActive'),
    totalClearedCard: document.getElementById('uploaderAgentCommissionTotalCleared'),
    tableBody: document.getElementById('agentCommissionTableBody'),
    modal: document.getElementById('uploaderAgentCommissionModal'),
    modalTitle: document.getElementById('uploaderAgentCommissionModalTitle'),
    modalSummary: document.getElementById('uploaderAgentCommissionModalSummary'),
    breakdownBody: document.getElementById('uploaderAgentCommissionBreakdownBody'),
    sentBtn: document.getElementById('uploaderAgentCommissionSentTabBtn'),
    activeBtn: document.getElementById('uploaderAgentCommissionActiveTabBtn'),
    clearedBtn: document.getElementById('uploaderAgentCommissionClearedTabBtn'),
    closeBtn: document.getElementById('closeUploaderAgentCommissionModal'),
    closeFooterBtn: document.getElementById('closeUploaderAgentCommissionModalFooterBtn')
  };
}

function ensureUploaderAgentCommissionUi() {
  const paidTab = document.getElementById('paidTab');
  if (!paidTab) return getUploaderAgentCommissionUiRefs();
  if (!document.getElementById('agentCommissionTableBody')) {
    paidTab.innerHTML = `
      <div class="agent-commission-summary uploader-agent-commission-overview" style="margin-bottom:16px;">
        <div class="agent-commission-summary-card commission-overview-card sent">
          <span class="agent-commission-summary-label">Total Sent to PFA</span>
          <strong id="uploaderAgentCommissionTotalSent">${formatCurrency(0)}</strong>
        </div>
        <div class="agent-commission-summary-card commission-overview-card active">
          <span class="agent-commission-summary-label">Total Commission Payable</span>
          <strong id="uploaderAgentCommissionTotalActive">${formatCurrency(0)}</strong>
        </div>
        <div class="agent-commission-summary-card commission-overview-card cleared">
          <span class="agent-commission-summary-label">Total Cleared</span>
          <strong id="uploaderAgentCommissionTotalCleared">${formatCurrency(0)}</strong>
        </div>
      </div>
      <div class="table-section" style="margin-top: 16px;">
        <div class="table-container">
          <table class="customers-table">
            <thead>
              <tr>
                <th>Agent</th>
                <th>Bank</th>
                <th>Account Number</th>
                <th>Total Commission</th>
                <th>View</th>
              </tr>
            </thead>
            <tbody id="agentCommissionTableBody">
              <tr><td colspan="5" class="loading-row"><div class="loading-spinner"></div> Loading agent commission records...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  if (!document.getElementById('uploaderAgentCommissionModal')) {
    const modal = document.createElement('div');
    modal.id = 'uploaderAgentCommissionModal';
    modal.className = 'modal';
    modal.innerHTML = `
      <div class="modal-content large-modal">
        <div class="modal-header">
          <h2 id="uploaderAgentCommissionModalTitle">Commission Breakdown</h2>
          <button class="close-btn" id="closeUploaderAgentCommissionModal">&times;</button>
        </div>
        <div class="modal-body">
          <div id="uploaderAgentCommissionModalSummary" class="agent-commission-summary"></div>
          <div class="subtab-strip agent-commission-subtabs">
            <button type="button" class="subtab-btn active" id="uploaderAgentCommissionSentTabBtn">Sent to PFA</button>
            <button type="button" class="subtab-btn" id="uploaderAgentCommissionActiveTabBtn">Commission Payable</button>
            <button type="button" class="subtab-btn" id="uploaderAgentCommissionClearedTabBtn">Cleared</button>
          </div>
          <div class="table-container">
            <table class="customers-table">
              <thead>
                <tr>
                  <th>Customer Name</th>
                  <th>25% Balance</th>
                  <th>Commission Amount</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody id="uploaderAgentCommissionBreakdownBody">
                <tr><td colspan="4" class="loading-row"><div class="loading-spinner"></div> Loading breakdown...</td></tr>
              </tbody>
            </table>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="cancel-btn" id="closeUploaderAgentCommissionModalFooterBtn">Close</button>
        </div>
      </div>
    `;
    document.body.appendChild(modal);
  }

  const refs = getUploaderAgentCommissionUiRefs();
  if (refs.sentBtn && !refs.sentBtn.dataset.bound) {
    refs.sentBtn.dataset.bound = 'true';
    refs.sentBtn.addEventListener('click', () => switchCommissionTab('sent_to_pfa'));
  }
  if (refs.activeBtn && !refs.activeBtn.dataset.bound) {
    refs.activeBtn.dataset.bound = 'true';
    refs.activeBtn.addEventListener('click', () => switchCommissionTab('active'));
  }
  if (refs.clearedBtn && !refs.clearedBtn.dataset.bound) {
    refs.clearedBtn.dataset.bound = 'true';
    refs.clearedBtn.addEventListener('click', () => switchCommissionTab('cleared'));
  }
  const closeModal = () => refs.modal?.classList.remove('active');
  if (refs.closeBtn && !refs.closeBtn.dataset.bound) {
    refs.closeBtn.dataset.bound = 'true';
    refs.closeBtn.addEventListener('click', closeModal);
  }
  if (refs.closeFooterBtn && !refs.closeFooterBtn.dataset.bound) {
    refs.closeFooterBtn.dataset.bound = 'true';
    refs.closeFooterBtn.addEventListener('click', closeModal);
  }
  if (refs.modal && !refs.modal.dataset.bound) {
    refs.modal.dataset.bound = 'true';
    refs.modal.addEventListener('click', (e) => {
      if (e.target === refs.modal) closeModal();
    });
  }
  return refs;
}

function isUploaderCommissionTrackable(sub) {
  const status = String(sub?.status || '').toLowerCase();
  return status === 'sent_to_pfa' || status === 'rsa_submitted' || status === 'paid' || status === 'cleared' || sub?.finalSubmitted === true || sub?.rsaSubmitted === true;
}

function getUploaderCommissionStatusBucket(sub) {
  const status = String(sub?.status || '').toLowerCase();
  if (status === 'cleared') return 'cleared';
  if (status === 'paid') return 'active';
  return 'sent_to_pfa';
}

function getUploaderAgentCommissionKey(subOrAgent) {
  const agentId = String(subOrAgent?.agentId || subOrAgent?.id || '').trim();
  return agentId ? `agent:${agentId}` : '';
}

function buildUploaderAgentCommissionGroups() {
  const groups = new Map();

  registeredAgents.forEach((agent) => {
    const key = getUploaderAgentCommissionKey(agent);
    groups.set(key, {
      key,
      agentName: String(agent?.fullName || 'No Agent').trim() || 'No Agent',
      bank: String(agent?.accountBank || '-').trim() || '-',
      accountNumber: String(agent?.accountNumber || '-').trim() || '-',
      totalCommission: 0,
      sentToPfaSubmissions: [],
      activeSubmissions: [],
      clearedSubmissions: [],
      sentToPfaCommission: 0,
      activeCommission: 0,
      clearedCommission: 0
    });
  });

  allSubmissions.forEach((sub) => {
    if (!isUploaderCommissionTrackable(sub)) return;
    if (!String(sub?.agentId || '').trim()) return;

    const key = getUploaderAgentCommissionKey(sub);
    if (!key) return;
    const snapshot = getSubmissionAgentSnapshot(sub);
    const group = groups.get(key) || {
      key,
      agentName: String(snapshot.agentName || 'No Agent').trim() || 'No Agent',
      bank: String(snapshot.agentAccountBank || '-').trim() || '-',
      accountNumber: String(snapshot.agentAccountNumber || '-').trim() || '-',
      totalCommission: 0,
      sentToPfaSubmissions: [],
      activeSubmissions: [],
      clearedSubmissions: [],
      sentToPfaCommission: 0,
      activeCommission: 0,
      clearedCommission: 0
    };
    const bucket = getUploaderCommissionStatusBucket(sub);
    const { commission2 } = getFinancials(sub);

    if (bucket === 'cleared') {
      group.clearedSubmissions.push(sub);
      group.clearedCommission += commission2;
    } else if (bucket === 'active') {
      group.activeSubmissions.push(sub);
      group.activeCommission += commission2;
    } else {
      group.sentToPfaSubmissions.push(sub);
      group.sentToPfaCommission += commission2;
    }
    group.totalCommission = group.sentToPfaCommission + group.activeCommission + group.clearedCommission;
    if (group.bank === '-' && String(snapshot.agentAccountBank || '').trim()) group.bank = String(snapshot.agentAccountBank).trim();
    if (group.accountNumber === '-' && String(snapshot.agentAccountNumber || '').trim()) group.accountNumber = String(snapshot.agentAccountNumber).trim();
    groups.set(key, group);
  });

  return Array.from(groups.values()).sort((a, b) => {
    if (b.totalCommission !== a.totalCommission) return b.totalCommission - a.totalCommission;
    return a.agentName.localeCompare(b.agentName);
  });
}

function renderUploaderAgentCommissionSummary(group) {
  const summary = document.getElementById('uploaderAgentCommissionModalSummary');
  if (!summary) return;
  summary.innerHTML = `
    <div class="agent-commission-summary-card">
      <span class="agent-commission-summary-label">Agent</span>
      <strong>${escapeHtml(group.agentName)}</strong>
    </div>
    <div class="agent-commission-summary-card">
      <span class="agent-commission-summary-label">Sent to PFA</span>
      <strong style="color:#2563eb;">${formatCurrency(group.sentToPfaCommission)}</strong>
    </div>
    <div class="agent-commission-summary-card">
      <span class="agent-commission-summary-label">Commission Payable</span>
      <strong style="color:#16a34a;">${formatCurrency(group.activeCommission)}</strong>
    </div>
    <div class="agent-commission-summary-card">
      <span class="agent-commission-summary-label">Cleared</span>
      <strong style="color:#dc2626;">${formatCurrency(group.clearedCommission)}</strong>
    </div>
    <div class="agent-commission-summary-card">
      <span class="agent-commission-summary-label">Total Commission</span>
      <strong>${formatCurrency(group.totalCommission)}</strong>
    </div>
  `;
}

function getCommissionAmountColor(bucket) {
  if (bucket === 'cleared') return '#dc2626';
  if (bucket === 'active') return '#16a34a';
  return '#2563eb';
}

function formatCommissionStatus(status) {
  const normalized = String(status || '').toLowerCase().trim();
  if (!normalized) return '-';
  if (normalized === 'sent_to_pfa' || normalized === 'rsa_submitted') return 'Sent to PFA';
  if (normalized === 'paid') return 'Paid';
  if (normalized === 'cleared') return 'Cleared';
  return normalized.replace(/_/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase());
}

function renderUploaderAgentCommissionBreakdown(view = 'sent_to_pfa') {
  const body = document.getElementById('uploaderAgentCommissionBreakdownBody');
  if (!body || !window.__currentUploaderAgentCommissionGroup) return;
  const group = window.__currentUploaderAgentCommissionGroup;
  const bucket = view === 'cleared' ? 'cleared' : view === 'active' ? 'active' : 'sent_to_pfa';
  const rows = (bucket === 'cleared' ? group.clearedSubmissions : bucket === 'active' ? group.activeSubmissions : group.sentToPfaSubmissions)
    .map((sub) => {
      const { twentyFive, commission2 } = getFinancials(sub);
      const amountColor = getCommissionAmountColor(bucket);
      return `
        <tr>
          <td><strong>${escapeHtml(sub.customerName || '-')}</strong></td>
          <td style="color:${amountColor};font-weight:700;">${formatCurrency(twentyFive)}</td>
          <td style="color:${amountColor};font-weight:700;">${formatCurrency(commission2)}</td>
          <td>${escapeHtml(formatCommissionStatus(sub.status || '-'))}</td>
        </tr>
      `;
    }).join('');
  const emptyLabel = bucket === 'sent_to_pfa' ? 'sent to PFA' : bucket === 'active' ? 'commission payable' : bucket;
  body.innerHTML = rows || `<tr><td colspan="4" class="no-data">No ${emptyLabel} records for this agent</td></tr>`;
}

window.openUploaderAgentCommissionModal = (groupKey) => {
  const refs = ensureUploaderAgentCommissionUi();
  const group = buildUploaderAgentCommissionGroups().find((entry) => entry.key === decodeURIComponent(String(groupKey || '').trim()));
  if (!group || !refs.modal) {
    showNotification('Agent commission breakdown not found', 'error');
    return;
  }
  window.__currentUploaderAgentCommissionGroup = group;
  refs.modalTitle.textContent = `${group.agentName} - Commission`;
  renderUploaderAgentCommissionSummary(group);
  switchCommissionTab('sent_to_pfa');
  refs.modal.classList.add('active');
};

function renderPaidTable() {
  const refs = ensureUploaderAgentCommissionUi();
  if (!refs.tableBody) return;

  const groups = buildUploaderAgentCommissionGroups();
  const submissionTotals = groups.reduce((acc, group) => {
    acc.sent += group.sentToPfaSubmissions.length;
    acc.active += group.activeSubmissions.length;
    acc.cleared += group.clearedSubmissions.length;
    acc.sentAmount += group.sentToPfaCommission;
    acc.activeAmount += group.activeCommission;
    acc.clearedAmount += group.clearedCommission;
    return acc;
  }, { sent: 0, active: 0, cleared: 0, sentAmount: 0, activeAmount: 0, clearedAmount: 0 });

  if (refs.totalSentCard) refs.totalSentCard.innerHTML = `${formatCurrency(submissionTotals.sentAmount)}<small>${submissionTotals.sent} app${submissionTotals.sent === 1 ? '' : 's'}</small>`;
  if (refs.totalActiveCard) refs.totalActiveCard.innerHTML = `${formatCurrency(submissionTotals.activeAmount)}<small>${submissionTotals.active} app${submissionTotals.active === 1 ? '' : 's'}</small>`;
  if (refs.totalClearedCard) refs.totalClearedCard.innerHTML = `${formatCurrency(submissionTotals.clearedAmount)}<small>${submissionTotals.cleared} app${submissionTotals.cleared === 1 ? '' : 's'}</small>`;

  if (paidCountBadge) {
    paidCountBadge.textContent = String(groups.length);
    paidCountBadge.style.display = 'inline-block';
  }

  if (!groups.length) {
    refs.tableBody.innerHTML = '<tr><td colspan="5" class="no-data">No commission records</td></tr>';
    return;
  }

  refs.tableBody.innerHTML = groups.map((group) => `
    <tr>
      <td><strong>${escapeHtml(group.agentName)}</strong></td>
      <td>${escapeHtml(group.bank)}</td>
      <td>${escapeHtml(group.accountNumber)}</td>
      <td><strong>${formatCurrency(group.totalCommission)}</strong></td>
      <td><button class="action-btn view-btn-small" onclick="window.openUploaderAgentCommissionModal('${encodeURIComponent(group.key)}')"><i class="fas fa-eye"></i> View</button></td>
    </tr>
  `).join('');
}

function showNotification(message, type = 'info') {
  if (type === 'error') {
    showErrorDetailModal(message);
    return;
  }
  if (!notification) return;
  notification.textContent = message;
  notification.className = `notification ${type}`;
  notification.style.display = 'block';
  setTimeout(() => { notification.style.display = 'none'; }, 3000);
}

async function loadApprovedAgents() {
  try {
    const currentUid = String(currentUser?.uid || '').trim();
    const currentEmail = String(currentUser?.email || '').trim().toLowerCase();
    let rows = [];

    if (currentUid) {
      const ownedSnap = await getDocs(query(
        collection(db, 'agents'),
        where('status', '==', 'approved'),
        where('createdByUid', '==', currentUid)
      ));
      rows = ownedSnap.docs.map((d) => ({ id: d.id, ...(d.data() || {}) }));
    } else {
      const snap = await getDocs(query(collection(db, 'agents'), where('status', '==', 'approved')));
      rows = snap.docs.map((d) => ({ id: d.id, ...(d.data() || {}) }));
    }

    approvedAgents = rows
      .filter((agent) => {
        if (currentUid && String(agent.createdByUid || '') === currentUid) return true;
        const createdByEmail = String(agent.createdBy || '').trim().toLowerCase();
        return !!currentEmail && createdByEmail === currentEmail;
      })
      .sort((a, b) => String(a.fullName || '').localeCompare(String(b.fullName || '')));
    populateApprovedAgentSelect();
  } catch (_) {
    approvedAgents = [];
    populateApprovedAgentSelect();
  }
}

async function openAgentRegistrationModal() {
  if (!agentRegistrationModal) return;
  await populateAgentBankOptions({ force: true });
  clearVerifiedAgentAccountLookup();
  setAgentAccountLookupStatus('', 'info');
  agentRegistrationModal.classList.add('active');
}

function closeAgentRegistrationModal() {
  if (!agentRegistrationModal) return;
  agentRegistrationModal.classList.remove('active');
  setAgentAccountLookupStatus('', 'info');
}

function normalizeAgentDate(value) {
  if (!value) return '-';
  if (value.toDate && typeof value.toDate === 'function') return safeFormatDate(value);
  if (value.seconds) return safeFormatDate(value);
  const asDate = new Date(value);
  if (Number.isNaN(asDate.getTime())) return '-';
  return safeFormatDate(asDate);
}

function statusBadgeForAgent(status) {
  const key = String(status || 'pending').toLowerCase();
  if (key === 'approved') return '<span class="status-badge status-approved">Approved</span>';
  if (key === 'rejected') return '<span class="status-badge status-rejected">Rejected</span>';
  return '<span class="status-badge status-pending">Pending</span>';
}

function escAttr(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function formatMoneyDisplay(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '-';
  const num = parseMoney(raw);
  if (!Number.isFinite(num) || num === 0) return raw === '0' ? formatCurrency(0) : raw;
  return formatCurrency(num);
}

function formatSubmissionStatusLabel(status) {
  const normalized = String(status || '').toLowerCase().trim();
  if (!normalized) return '-';
  if (normalized === 'processing_to_pfa' || normalized === 'approved') return 'Processing to PFA';
  if (normalized === 'sent_to_pfa' || normalized === 'rsa_submitted') return 'Sent to PFA';
  if (normalized === 'rejected_by_rsa') return 'Rejected by RSA';
  if (normalized === 'paid') return 'Paid';
  if (normalized === 'cleared') return 'Cleared';
  return normalized.replace(/_/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase());
}

function getSubmissionDetailValue(submission = {}, keys = [], fallback = '-') {
  const details = submission?.customerDetails || {};
  for (const key of keys) {
    const detailValue = details?.[key];
    if (detailValue !== undefined && detailValue !== null && String(detailValue).trim() !== '') return String(detailValue);
    const rootValue = submission?.[key];
    if (rootValue !== undefined && rootValue !== null && String(rootValue).trim() !== '') return String(rootValue);
  }
  return fallback;
}

window.viewSubmissionDetails = (submissionId) => {
  const sub = allSubmissions.find((item) => item.id === submissionId);
  if (!sub) {
    showNotification('Submission details not found', 'error');
    return;
  }
  const agentSnapshot = getSubmissionAgentSnapshot(sub);
  const detailSections = [
    {
      title: 'Customer Details',
      fields: [
        ['Customer Name', getSubmissionDetailValue(sub, ['name', 'customerName'], sub.customerName || '-')],
        ['Date of Birth', getSubmissionDetailValue(sub, ['dob', 'dateOfBirth', 'customerDob'])],
        ['Email', getSubmissionDetailValue(sub, ['email', 'customerEmail'])],
        ['Phone', getSubmissionDetailValue(sub, ['phone', 'customerPhone'])],
        ['NIN', getSubmissionDetailValue(sub, ['nin', 'customerNIN'])],
        ['Address', getSubmissionDetailValue(sub, ['address', 'customerAddress'])],
        ['Account Number', getSubmissionDetailValue(sub, ['accountNo'])],
        ['Employer', getSubmissionDetailValue(sub, ['employer'])],
        ['Originating TP', getSubmissionDetailValue(sub, ['originatingTP'])]
      ]
    },
    {
      title: 'PFA Details',
      fields: [
        ['PFA', getSubmissionDetailValue(sub, ['pfa', 'pfaName'])],
        ['PEN Number', getSubmissionDetailValue(sub, ['penNo'])],
        ['RSA Statement Date', getSubmissionDetailValue(sub, ['rsaStatementDate'])],
        ['RSA Balance', formatMoneyDisplay(getSubmissionDetailValue(sub, ['rsaBalance'], ''))],
        ['25% RSA', formatMoneyDisplay(getSubmissionDetailValue(sub, ['rsa25', 'rsa25Percent'], ''))]
      ]
    },
    {
      title: 'Property Details',
      fields: [
        ['Property Type', getSubmissionDetailValue(sub, ['propertyType'])],
        ['House Number', getSubmissionDetailValue(sub, ['houseNumber'])],
        ['Property Value', formatMoneyDisplay(getSubmissionDetailValue(sub, ['propertyValue'], ''))],
        ['Facility Fee', formatMoneyDisplay(getSubmissionDetailValue(sub, ['facilityFee'], ''))],
        ['Loan Amount', formatMoneyDisplay(getSubmissionDetailValue(sub, ['loanAmount'], ''))],
        ['Tenor', getSubmissionDetailValue(sub, ['tenor'])]
      ]
    },
    {
      title: 'Agent Details',
      fields: [
        ['Agent Name', String(agentSnapshot.agentName || '').trim() || 'No Agent'],
        ['Agent Contact', String(agentSnapshot.agentContactNumber || '').trim() || '-'],
        ['Agent Account Number', String(agentSnapshot.agentAccountNumber || '').trim() || '-'],
        ['Agent Account Bank', String(agentSnapshot.agentAccountBank || '').trim() || '-']
      ]
    }
  ];

  const detailsModal = document.createElement('div');
  detailsModal.className = 'modal active';
  detailsModal.innerHTML = `
    <div class="modal-content large-modal">
      <div class="modal-header">
        <h2><i class="fas fa-circle-info"></i> Application Details</h2>
        <button class="close-btn" type="button">&times;</button>
      </div>
      <div class="modal-body">
        <div class="document-info" style="margin-bottom:18px;">
          <h3>${escapeHtml(sub.customerName || 'Unknown Customer')}</h3>
          <p><i class="fas fa-hashtag"></i> Application ID: ${escapeHtml(sub.id)}</p>
          <p><i class="fas fa-user-check"></i> Status: ${escapeHtml(formatSubmissionStatusLabel(sub.status || 'pending'))}</p>
        </div>
        ${detailSections.map((section) => `
          <div class="form-section" style="margin-bottom:16px;">
            <div class="section-header">
              <h3>${escapeHtml(section.title)}</h3>
            </div>
            <div class="section-body">
              <div class="customer-input-grid">
                ${section.fields.map(([label, value]) => `
                  <div>
                    <label>${escapeHtml(label)}</label>
                    <input type="text" readonly value="${escAttr(value)}">
                  </div>
                `).join('')}
              </div>
            </div>
          </div>
        `).join('')}
      </div>
      <div class="modal-footer">
        <button class="cancel-btn" type="button">Close</button>
      </div>
    </div>
  `;
  document.body.appendChild(detailsModal);
  const close = () => detailsModal.remove();
  detailsModal.querySelectorAll('button').forEach((btn) => btn.addEventListener('click', close));
  detailsModal.addEventListener('click', (event) => {
    if (event.target === detailsModal) close();
  });
};

window.viewAgentDetails = async (agentId) => {
  let agent =
    registeredAgents.find((row) => row.id === agentId) ||
    approvedAgents.find((row) => row.id === agentId) ||
    null;

  if (!agent) {
    try {
      const snap = await getDoc(doc(db, 'agents', agentId));
      if (snap.exists()) {
        agent = { id: snap.id, ...(snap.data() || {}) };
      }
    } catch (_) {
    }
  }

  if (!agent) {
    showNotification('Agent record not found', 'error');
    return;
  }

  const detailsModal = document.createElement('div');
  detailsModal.className = 'modal active';
  detailsModal.innerHTML = `
    <div class="modal-content">
      <div class="modal-header">
        <h2><i class="fas fa-user-tie"></i> Agent Details</h2>
        <button class="close-btn" type="button">&times;</button>
      </div>
      <div class="modal-body">
        <div class="customer-input-grid">
          <div><label>Full Name</label><input type="text" readonly value="${escAttr(agent.fullName || '-')}"></div>
          <div><label>Contact Number</label><input type="text" readonly value="${escAttr(agent.contactNumber || '-')}"></div>
          <div><label>Account Number</label><input type="text" readonly value="${escAttr(agent.accountNumber || '-')}"></div>
          <div><label>Account Bank</label><input type="text" readonly value="${escAttr(agent.accountBank || '-')}"></div>
          <div><label>Status</label><input type="text" readonly value="${escAttr(String(agent.status || 'pending'))}"></div>
          <div><label>Date Registered</label><input type="text" readonly value="${escAttr(normalizeAgentDate(agent.createdAt))}"></div>
        </div>
      </div>
      <div class="modal-footer">
        <button class="cancel-btn" type="button">Close</button>
      </div>
    </div>
  `;
  document.body.appendChild(detailsModal);
  const close = () => detailsModal.remove();
  detailsModal.querySelectorAll('button').forEach((btn) => btn.addEventListener('click', close));
  detailsModal.addEventListener('click', (e) => {
    if (e.target === detailsModal) close();
  });
};

async function loadRegisteredAgents() {
  if (!registeredAgentsTableBody) return;
  registeredAgentsTableBody.innerHTML = '<tr><td colspan="7" class="loading-row"><div class="loading-spinner"></div> Loading agents...</td></tr>';
  try {
    const currentUid = String(currentUser?.uid || '').trim();
    const currentEmail = String(currentUser?.email || '').trim().toLowerCase();
    const byId = new Map();

    if (currentUid) {
      const ownSnap = await getDocs(query(collection(db, 'agents'), where('createdByUid', '==', currentUid)));
      ownSnap.docs.forEach((row) => {
        byId.set(row.id, { id: row.id, ...(row.data() || {}) });
      });
    }
    if (currentEmail) {
      const emailSnap = await getDocs(query(collection(db, 'agents'), where('createdBy', '==', currentEmail)));
      emailSnap.docs.forEach((row) => {
        if (!byId.has(row.id)) byId.set(row.id, { id: row.id, ...(row.data() || {}) });
      });
    }

    registeredAgents = Array.from(byId.values()).sort((a, b) => {
      const aTime = getDateFromAny(a.createdAt)?.getTime() || 0;
      const bTime = getDateFromAny(b.createdAt)?.getTime() || 0;
      return bTime - aTime;
    });

    renderPaidTable();

    if (!registeredAgents.length) {
      registeredAgentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No registered agents yet</td></tr>';
      return;
    }

    registeredAgentsTableBody.innerHTML = registeredAgents.map((agent) => `
      <tr>
        <td><strong>${agent.fullName || '-'}</strong></td>
        <td>${agent.contactNumber || '-'}</td>
        <td>${agent.accountNumber || '-'}</td>
        <td>${agent.accountBank || '-'}</td>
        <td>${statusBadgeForAgent(agent.status)}</td>
        <td>${normalizeAgentDate(agent.createdAt)}</td>
        <td><button class="action-btn view-btn-small" onclick="window.viewAgentDetails('${agent.id}')"><i class="fas fa-eye"></i> View</button></td>
      </tr>
    `).join('');
  } catch (error) {
    registeredAgents = [];
    renderPaidTable();
    registeredAgentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">Unable to load agents</td></tr>';
  }
}

function populateApprovedAgentSelect() {
  if (!customerAgentSelect) return;
  const current = String(customerAgentSelect.value || '');
  customerAgentSelect.innerHTML = '<option value="">No Agent</option>' + approvedAgents.map((agent) => (
    `<option value="${agent.id}">${agent.fullName || 'Unnamed'} - ${agent.accountNumber || '-'}</option>`
  )).join('');
  if (current && approvedAgents.some((a) => a.id === current)) {
    customerAgentSelect.value = current;
    return;
  }
  if (currentSubmissionAgentFallback && !approvedAgents.some((a) => a.id === currentSubmissionAgentFallback.value)) {
    const labelName = currentSubmissionAgentFallback.fullName || 'Linked Agent';
    const labelAccountNumber = currentSubmissionAgentFallback.accountNumber || '-';
    const opt = document.createElement('option');
    opt.value = currentSubmissionAgentFallback.value;
    opt.textContent = `${labelName} - ${labelAccountNumber} (linked)`;
    customerAgentSelect.appendChild(opt);
    if (current === currentSubmissionAgentFallback.value) {
      customerAgentSelect.value = currentSubmissionAgentFallback.value;
    }
  }
}

async function handleAgentRegistration(e) {
  e.preventDefault();
  const contactNumber = String(document.getElementById('agentContactNumber')?.value || '').trim();
  const accountNumber = String(agentAccountNumberInput?.value || '').replace(/\D/g, '');
  const selectedBank = getSelectedAgentAccountBank();
  const accountBank = selectedBank.name;
  const accountBankCode = selectedBank.code;
  const accountName = String(agentAccountNameInput?.value || verifiedAgentAccountLookup.accountName || '').trim();

  if (!contactNumber || !accountNumber || !accountBank) {
    showNotification('Please complete all agent fields', 'error');
    return;
  }
  if (!accountBankCode) {
    showNotification('Select a valid account bank from the list', 'error');
    return;
  }
  if (!/^\d{10,11}$/.test(contactNumber.replace(/\D/g, ''))) {
    showNotification('Agent contact number must be 10 or 11 digits', 'error');
    return;
  }
  if (!/^\d{10}$/.test(accountNumber.replace(/\D/g, ''))) {
    showNotification('Agent account number must be exactly 10 digits', 'error');
    return;
  }
  if (
    verifiedAgentAccountLookup.accountNumber !== accountNumber ||
    verifiedAgentAccountLookup.bankCode !== accountBankCode ||
    !accountName
  ) {
    const verified = await resolveAgentAccountName({ silent: false });
    if (!verified || !String(verifiedAgentAccountLookup.accountName || '').trim()) {
      showNotification('Verify the agent account name before submitting.', 'error');
      return;
    }
  }
  const finalAccountName = String(verifiedAgentAccountLookup.accountName || accountName).trim();
  if (!finalAccountName) {
    showNotification('Verify the agent account name before submitting.', 'error');
    return;
  }
  const fullName = finalAccountName;

  try {
    if (submitAgentFormBtn) {
      submitAgentFormBtn.disabled = true;
      submitAgentFormBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Submitting...';
    }
    const systemSettings = await getSystemSettings(db, { force: true });
    const approvalRequired = systemSettings.agentRegistrationRules?.approvalRequired !== false;
    const agentRef = await addDoc(collection(db, 'agents'), {
      fullName,
      contactNumber,
      accountNumber,
      accountBank,
      accountBankCode,
      accountName: finalAccountName,
      status: approvalRequired ? 'pending' : 'approved',
      createdBy: currentUser?.email || '',
      createdByUid: currentUser?.uid || '',
      createdAt: serverTimestamp()
    });
    await notifyAdminPushEvent({
      currentUser,
      eventType: 'new_agent_registration',
      title: approvalRequired ? 'New Agent Registration' : 'Agent Registered',
      body: approvalRequired
        ? `${fullName} was submitted for admin approval.`
        : `${fullName} was registered successfully.`,
      clickUrl: '/admin-dashboard.html',
      meta: {
        agentId: agentRef.id,
        fullName,
        contactNumber,
        accountNumber,
        accountBank,
        accountBankCode,
        accountName: finalAccountName,
        createdBy: currentUser?.email || '',
        approvalRequired
      }
    });
    showNotification(approvalRequired ? 'Agent registration submitted for admin approval' : 'Agent registered successfully', 'success');
    agentRegistrationForm?.reset();
    clearVerifiedAgentAccountLookup();
    setAgentAccountLookupStatus('', 'info');
    closeAgentRegistrationModal();
    await loadRegisteredAgents();
  } catch (err) {
    showNotification('Failed to submit agent registration', 'error');
  } finally {
    if (submitAgentFormBtn) {
      submitAgentFormBtn.disabled = false;
      submitAgentFormBtn.innerHTML = '<i class="fas fa-paper-plane"></i> Submit for Approval';
    }
  }
}

function getFileIcon(filename) {
  const ext = filename?.split('.').pop().toLowerCase();
  const icons = { 'pdf': 'fa-file-pdf', 'doc': 'fa-file-word', 'docx': 'fa-file-word', 'jpg': 'fa-file-image', 'jpeg': 'fa-file-image', 'png': 'fa-file-image' };
  return icons[ext] || 'fa-file-alt';
}

function formatFileSize(bytes) {
  if (!bytes) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function hideLoader() {
  const loader = document.getElementById('globalLoader');
  if (loader) loader.style.display = 'none';
}
window.hideLoader = hideLoader;
window.showLoader = function(message = 'Processing...') {
  const loader = document.getElementById('globalLoader');
  const loaderText = document.getElementById('loaderText');
  if (loader) loader.style.display = 'flex';
  if (loaderText) loaderText.textContent = message;
};

// ==================== MAKE FUNCTIONS GLOBAL ====================
if (typeof window.openSingleUpload !== 'function') window.openSingleUpload = openSingleUpload;
if (typeof window.viewDocument !== 'function') window.viewDocument = viewDocument;
if (typeof window.viewSubmissionDocs !== 'function') window.viewSubmissionDocs = viewSubmissionDocs;
if (typeof window.openEditModal !== 'function') window.openEditModal = openEditModal;
if (typeof window.switchTab !== 'function') window.switchTab = switchTab;
if (typeof window.showApplicationTrack !== 'function') window.showApplicationTrack = showApplicationTrack;
if (typeof window.saveCustomerDetails !== 'function') window.saveCustomerDetails = saveCustomerDetails;
if (typeof window.resetCustomerDetails !== 'function') window.resetCustomerDetails = resetCustomerDetails;

window.removeUploadedDoc = (docType) => {
  if (!currentCustomerUploads[docType] || currentCustomerUploads[docType].length === 0) return;
  currentCustomerUploads[docType].pop();
  renderDocumentGrid(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'upload');
  renderDocumentGrid(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'upload');
  updateSubmitButton();
  persistDraftSilentlyIfNeeded();
};

// ==================== INITIALIZE RECENT SEARCH ====================
document.addEventListener('DOMContentLoaded', function() {
  const recentSearch = document.getElementById('recentSearch');
  if (recentSearch) {
    recentSearch.addEventListener('input', function(e) {
      const searchTerm = e.target.value.toLowerCase();
      const rows = document.querySelectorAll('#recentTableBody tr');
      rows.forEach(row => { const text = row.textContent.toLowerCase(); row.style.display = text.includes(searchTerm) ? '' : 'none'; });
    });
  }
  ['draft', 'pending', 'approved', 'rejected'].forEach(tab => {
    const searchInput = document.getElementById(`${tab}Search`);
    if (searchInput) {
      searchInput.addEventListener('input', function(e) {
        const searchTerm = e.target.value.toLowerCase();
        const rows = document.querySelectorAll(`#${tab}TableBody tr`);
        rows.forEach(row => { const text = row.textContent.toLowerCase(); row.style.display = text.includes(searchTerm) ? '' : 'none'; });
      });
    }
  });
});
