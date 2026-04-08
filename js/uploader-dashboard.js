// js/uploader-dashboard.js - LIGHTWEIGHT VERSION (Only Track + Calculator)
import { auth, db } from './firebase-config.js';
import {
    collection, query, where, orderBy, onSnapshot, getDocs, getDoc, doc, limit
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";

let currentUser = null;
let allSubmissions = [];
const userFullNames = new Map();

// ==================== PROPERTY RULES (RSA Balance Ranges) ====================
const PROPERTY_RULES = [
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

function determinePropertyByRsa(rsaAmount) {
    const n = Number(rsaAmount) || 0;
    for (const r of PROPERTY_RULES) {
        if (n >= r.min && n <= r.max) return r;
    }
    return null;
}

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (e) {
        return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
}

// DOM Elements
const userName = document.getElementById('userName');
const userAvatar = document.getElementById('userAvatar');
const pageTitle = document.getElementById('pageTitle');

// ==================== Track Modal Elements ====================
const globalTrackBtn = document.getElementById('globalTrackBtn');
const globalTrackModal = document.getElementById('globalTrackModal');
const closeGlobalTrackModal = document.getElementById('closeGlobalTrackModal');
const cancelGlobalTrack = document.getElementById('cancelGlobalTrack');
const searchTrackBtn = document.getElementById('searchTrackBtn');
const trackSearchType = document.getElementById('trackSearchType');
const trackSearchInput = document.getElementById('trackSearchInput');
const trackSearchResults = document.getElementById('trackSearchResults');
const trackResultsList = document.getElementById('trackResultsList');

// ==================== Loan Calculator Elements ====================
const loanCalcBtn = document.getElementById('loanCalcBtn');
const loanCalcModal = document.getElementById('loanCalculatorModal');
const closeLoanCalcModal = document.getElementById('closeLoanCalcModal');
const closeLoanCalcBtn = document.getElementById('closeLoanCalcBtn');
const calculateLoanBtn = document.getElementById('calculateLoanBtn');
const calcRsaBalance = document.getElementById('calcRsaBalance');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    auth.onAuthStateChanged(async (user) => {
        if (user) {
            currentUser = user;
            try {
                const q = query(collection(db, 'users'), where('email', '==', user.email));
                const snap = await getDocs(q);
                if (!snap.empty) {
                    const data = snap.docs[0].data();
                    userName.textContent = data.fullName || user.displayName || user.email.split('@')[0];
                } else {
                    userName.textContent = user.displayName || user.email.split('@')[0];
                }
            } catch (e) {
                userName.textContent = user.displayName || user.email.split('@')[0];
            }
            userAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3E👤%3C/text%3E%3C/svg%3E';

            // Just load submissions for display - don't interfere with uploader
            loadSubmissionsForDisplay();
        } else {
            window.location.href = 'index.html';
        }
    });

    setupTrackListeners();
    setupLoanCalculator();
    setupTabSwitching();
    setupForceRefreshButtons();
    setupIdleLogout();
    ensureSignOutUser();
});

function ensureSignOutUser() {
    try {
        const desc = Object.getOwnPropertyDescriptor(window, 'signOutUser');
        // `js/auth.js` defines a non-writable `signOutUser`; reuse it when present.
        if (desc && desc.writable === false) return;
        window.signOutUser = async () => {
            try { await signOut(auth); } catch (e) { }
            window.location.href = 'index.html';
        };
    } catch (e) { /* ignore */ }
}

const IDLE_TIMEOUT_MS = 30 * 60 * 1000;
let idleLastActivity = Date.now();
let idleIntervalHandle = null;

function setupIdleLogout() {
    if (idleIntervalHandle) return;
    const bump = () => { idleLastActivity = Date.now(); };
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach((ev) => {
        window.addEventListener(ev, bump, { passive: true });
    });
    window.addEventListener('focus', bump);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) bump(); });

    idleIntervalHandle = setInterval(() => {
        if (!auth.currentUser) return;
        if (Date.now() - idleLastActivity >= IDLE_TIMEOUT_MS) {
            window.signOutUser();
        }
    }, 60 * 1000);
}

function forceHardRefresh() {
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
}

function setupForceRefreshButtons() {
    document.getElementById('forceRefreshBtn')?.addEventListener('click', (e) => {
        e.preventDefault();
        forceHardRefresh();
    });
    document.getElementById('forceRefreshBtnMobile')?.addEventListener('click', (e) => {
        e.preventDefault();
        forceHardRefresh();
    });
}

function setupTabSwitching() {
    document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const tabId = item.dataset.tab;

            document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
            item.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            const targetTab = document.getElementById(`${tabId}Tab`);
            if(targetTab) targetTab.classList.add('active');

            const titles = {
                pending: 'Pending Review',
                overview: 'Dashboard',
                approved: 'Approved Documents',
                rejected: 'Rejected Submissions'
            };
            if (pageTitle) pageTitle.textContent = titles[tabId];
        });
    });
}

function setupTrackListeners() {
    if (globalTrackBtn && globalTrackModal) {
        globalTrackBtn.addEventListener('click', () => {
            if (globalTrackModal) {
                trackSearchInput.value = '';
                trackSearchResults.style.display = 'none';
                trackResultsList.innerHTML = '';
                globalTrackModal.classList.add('active');
            }
        });
    }

    if (closeGlobalTrackModal) {
        closeGlobalTrackModal.addEventListener('click', () => {
            globalTrackModal.classList.remove('active');
        });
    }

    if (cancelGlobalTrack) {
        cancelGlobalTrack.addEventListener('click', () => {
            globalTrackModal.classList.remove('active');
        });
    }

    if (searchTrackBtn) {
        searchTrackBtn.addEventListener('click', performGlobalTrackSearch);
    }

    if (trackSearchInput) {
        trackSearchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                performGlobalTrackSearch();
            }
        });
    }
}

function setupLoanCalculator() {
    if (loanCalcBtn && loanCalcModal) {
        loanCalcBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            if (calcRsaBalance) calcRsaBalance.value = '';
            document.getElementById('calc25Percent').textContent = '₦0.00';
            document.getElementById('calcFacilityFee').textContent = '₦0.00';
            document.getElementById('calcPropertyType').textContent = '-';
            document.getElementById('calcPropertyValue').textContent = '₦0.00';
            document.getElementById('calcLoanAmount').textContent = '₦0.00';
            const calcResults = document.getElementById('calcResults');
            if (calcResults) calcResults.style.display = 'none';
            loanCalcModal.classList.add('active');
        });
    }

    if (closeLoanCalcModal) {
        closeLoanCalcModal.addEventListener('click', function() {
            loanCalcModal.classList.remove('active');
        });
    }

    if (closeLoanCalcBtn) {
        closeLoanCalcBtn.addEventListener('click', function() {
            loanCalcModal.classList.remove('active');
        });
    }

    if (calculateLoanBtn) {
        calculateLoanBtn.addEventListener('click', function(e) {
            e.preventDefault();
            calculateLoan();
        });
    }

    if (calcRsaBalance) {
        calcRsaBalance.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                calculateLoan();
            }
        });
        calcRsaBalance.addEventListener('input', function(e) {
            // Allow digits and a single decimal point.
            let value = String(e.target.value || '').replace(/[^0-9.]/g, '');
            const firstDot = value.indexOf('.');
            if (firstDot !== -1) {
                value = value.slice(0, firstDot + 1) + value.slice(firstDot + 1).replace(/\./g, '');
            }
            e.target.value = value;
        });
    }
}

function calculateLoan() {
    if (!calcRsaBalance) return;
    const balance = parseFloat(calcRsaBalance.value) || 0;
    if (!balance || balance < 4000000) {
        showNotification('Please enter a valid RSA balance (minimum ₦4,000,000)', 'warning');
        return;
    }
    if (balance > 299999999) {
        showNotification('RSA balance exceeds maximum allowed (₦299,999,999)', 'error');
        return;
    }
    const rule = determinePropertyByRsa(balance);
    if (!rule) {
        showNotification('RSA balance outside property range', 'error');
        return;
    }
    const twentyFivePercent = balance * 0.25;
    // Round up to nearest hundred.
    const loanAmount = Math.ceil((rule.value - twentyFivePercent) / 100) * 100;
    document.getElementById('calc25Percent').textContent = formatCurrency(twentyFivePercent);
    document.getElementById('calcFacilityFee').textContent = formatCurrency(rule.fee);
    document.getElementById('calcPropertyType').textContent = rule.name;
    document.getElementById('calcPropertyValue').textContent = formatCurrency(rule.value);
    document.getElementById('calcLoanAmount').textContent = formatCurrency(loanAmount);
    const calcResults = document.getElementById('calcResults');
    if (calcResults) calcResults.style.display = 'block';
}

async function performGlobalTrackSearch() {
    const searchType = trackSearchType.value;
    const searchTerm = trackSearchInput.value.trim();
    if (!searchTerm) {
        showNotification('Please enter a search term', 'error');
        return;
    }
    showLoader('Searching applications...');
    try {
        const submissionsRef = collection(db, 'submissions');
        let results = [];
        if (searchType === 'id') {
            const docRef = doc(db, 'submissions', searchTerm);
            const docSnap = await getDoc(docRef);
            if (docSnap.exists()) {
                results = [{ id: docSnap.id, ...docSnap.data() }];
            }
        } else {
            const allSubmissionsQuery = query(submissionsRef, orderBy('customerName'), limit(100));
            const querySnapshot = await getDocs(allSubmissionsQuery);
            querySnapshot.forEach((doc) => {
                const data = doc.data();
                const customerName = (data.customerName || '').toLowerCase();
                if (customerName.includes(searchTerm.toLowerCase())) {
                    results.push({ id: doc.id, ...data });
                }
            });
        }
        results.sort((a, b) => (a.customerName || '').localeCompare(b.customerName || ''));
        displayTrackResults(results);
    } catch (error) {
        showNotification('Error searching applications: ' + error.message, 'error');
    } finally {
        hideLoader();
    }
}

async function displayTrackResults(results) {
    trackResultsList.innerHTML = '';
    if (results.length === 0) {
        trackResultsList.innerHTML = '<div style="text-align: center; padding: 30px; color: #64748b;">No applications found matching your search.</div>';
        trackSearchResults.style.display = 'block';
        return;
    }
    const resultCount = document.createElement('div');
    resultCount.style.cssText = 'padding: 10px; background: #f8fafc; border-radius: 8px; margin-bottom: 15px; font-weight: 600;';
    resultCount.innerHTML = `<strong>${results.length}</strong> application(s) found`;
    trackResultsList.appendChild(resultCount);

    for (const sub of results) {
        const resultItem = document.createElement('div');
        resultItem.className = 'result-item';
        resultItem.style.cssText = 'padding: 15px; border: 1px solid #e2e8f0; border-radius: 8px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s;';
        resultItem.onmouseover = () => { resultItem.style.backgroundColor = '#f8fafc'; };
        resultItem.onmouseout = () => { resultItem.style.backgroundColor = 'white'; };
        resultItem.onclick = () => {
            globalTrackModal.classList.remove('active');
            // Use the existing showApplicationTrack from document-uploader.js
            if (typeof window.showApplicationTrack === 'function') {
                window.showApplicationTrack(sub.id);
            } else {
                alert('Application ID: ' + sub.id);
            }
        };

        const uploadedByName = sub.uploadedBy ? await getUserFullName(sub.uploadedBy) : 'Unknown';
        resultItem.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <strong style="font-size: 16px; color: #003366;">${sub.customerName || 'Unknown'}</strong>
                <span style="padding: 4px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; background: ${
                    sub.status === 'approved' ? '#d1fae5' :
                    sub.status === 'rejected' ? '#fee2e2' : '#fef3c7'
                }; color: ${
                    sub.status === 'approved' ? '#065f46' :
                    sub.status === 'rejected' ? '#991b1b' : '#92400e'
                };">${sub.status || 'pending'}</span>
            </div>
            <div style="display: flex; gap: 15px; font-size: 12px; color: #64748b; margin-bottom: 8px; flex-wrap: wrap;">
                <span><i class="fas fa-id-card"></i> ID: ${sub.id.substring(0, 8)}...</span>
                <span><i class="fas fa-calendar"></i> ${safeFormatDate(sub.uploadedAt)}</span>
                ${sub.uploadedBy ? `<span><i class="fas fa-user"></i> Uploader: ${uploadedByName}</span>` : ''}
            </div>
        `;
        trackResultsList.appendChild(resultItem);
    }
    trackSearchResults.style.display = 'block';
}

// ==================== GET USER FULL NAME BY EMAIL ====================
async function getUserFullName(email) {
    if (!email) return 'Unknown';
    if (userFullNames.has(email)) return userFullNames.get(email);
    try {
        const userQuery = query(collection(db, 'users'), where('email', '==', email));
        const userSnapshot = await getDocs(userQuery);
        if (!userSnapshot.empty) {
            const userData = userSnapshot.docs[0].data();
            const fullName = userData.fullName || userData.displayName || email.split('@')[0];
            userFullNames.set(email, fullName);
            return fullName;
        }
    } catch (err) {
        // Silently fail
    }
    const fallbackName = email.split('@')[0];
    userFullNames.set(email, fallbackName);
    return fallbackName;
}

// ==================== LOAD SUBMISSIONS FOR DISPLAY ONLY ====================
function loadSubmissionsForDisplay() {
    if (!currentUser) return;

    const q = query(
        collection(db, 'submissions'),
        where('uploadedBy', '==', currentUser.email),
        orderBy('uploadedAt', 'desc')
    );

    onSnapshot(q, async (snapshot) => {
        allSubmissions = [];
        const emails = new Set();
        snapshot.forEach((doc) => {
            const data = doc.data();
            allSubmissions.push({ id: doc.id, ...data });
            if (data.uploadedBy) emails.add(data.uploadedBy);
            if (data.reviewedBy) emails.add(data.reviewedBy);
            if (data.assignedTo) emails.add(data.assignedTo);
        });

        await ensureUserFullNames(Array.from(emails));

        // Update dashboard counts only - don't render tables that conflict
        updateDashboardCards();
        renderRecentTable();
    }, (error) => {
        // Silently fail
    });
}

async function ensureUserFullNames(emails) {
    if (!emails || emails.length === 0) return;
    for (const email of emails) {
        if (!email || userFullNames.has(email)) continue;
        try {
            const q = query(collection(db, 'users'), where('email', '==', email));
            const snap = await getDocs(q);
            if (!snap.empty) {
                const d = snap.docs[0].data();
                userFullNames.set(email, d.fullName || d.displayName || email.split('@')[0]);
            } else {
                userFullNames.set(email, email.split('@')[0]);
            }
        } catch (e) {
            userFullNames.set(email, email.split('@')[0]);
        }
    }
}

function safeFormatDate(dateValue) {
    if (!dateValue) return 'N/A';
    try {
        let date;
        if (dateValue.toDate && typeof dateValue.toDate === 'function') date = dateValue.toDate();
        else if (typeof dateValue === 'string') date = new Date(dateValue);
        else if (dateValue.seconds) date = new Date(dateValue.seconds * 1000);
        else if (dateValue instanceof Date) date = dateValue;
        else return 'N/A';
        return date.toLocaleDateString('en-NG', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    } catch (error) {
        return 'Invalid date';
    }
}

function updateDashboardCards() {
    const pending = allSubmissions.filter(s => s.status === 'pending').length;
    const approved = allSubmissions.filter(s => s.status === 'approved').length;
    const rejected = allSubmissions.filter(s => s.status === 'rejected').length;
    document.getElementById('cardPendingCount') && (document.getElementById('cardPendingCount').textContent = pending);
    document.getElementById('cardApprovedCount') && (document.getElementById('cardApprovedCount').textContent = approved);
    document.getElementById('cardRejectedCount') && (document.getElementById('cardRejectedCount').textContent = rejected);
}

function renderRecentTable() {
    const rows = document.getElementById('recentTableBody');
    if (!rows) return;

    const recent = allSubmissions.slice(0, 10);

    if (recent.length === 0) {
        rows.innerHTML = '<tr><td colspan="5" class="no-data">No recent applications</td></tr>';
        return;
    }

    rows.innerHTML = recent.map(sub => {
        const uAt = sub.uploadedAt ? (sub.uploadedAt.toDate ? sub.uploadedAt.toDate() : new Date(sub.uploadedAt)) : null;
        const uploadDate = uAt ? uAt.toLocaleString('en-NG', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' }) : 'N/A';
        const uploaderName = (sub.uploadedBy && userFullNames.get(sub.uploadedBy)) ? userFullNames.get(sub.uploadedBy) : (sub.uploadedBy || '-');
        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td><span class="status-badge status-${sub.status}" style="padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600; background:${
                    sub.status === 'approved' ? '#d1fae5' :
                    sub.status === 'rejected' ? '#fee2e2' : '#fef3c7'
                }; color:${
                    sub.status === 'approved' ? '#065f46' :
                    sub.status === 'rejected' ? '#991b1b' : '#92400e'
                };">${sub.status ? sub.status.charAt(0).toUpperCase() + sub.status.slice(1) : 'Pending'}</span></td>
                <td>${uploaderName}</td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i></button>
                    <button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')" style="margin-left:5px; background:#8b5cf6; color:white; border:none; padding:5px 10px; border-radius:4px;"><i class="fas fa-map-marker-alt"></i></button>
                </td>
            </tr>
        `;
    }).join('');
}

// Simple notification
function showNotification(msg, type = 'info') {
    const el = document.getElementById('notification');
    if(!el) return;
    el.textContent = msg;
    el.className = `notification ${type}`;
    el.style.display = 'block';
    if (type === 'success') el.style.background = '#10b981';
    else if (type === 'error') el.style.background = '#ef4444';
    else if (type === 'warning') el.style.background = '#f59e0b';
    else el.style.background = '#3b82f6';
    el.style.color = 'white';
    setTimeout(() => { el.style.display = 'none'; }, 3000);
}

// Loader Functions
function showLoader(msg) {
    const loader = document.getElementById('globalLoader');
    const text = document.getElementById('loaderText');
    if(loader && text) {
        text.textContent = msg || "Processing...";
        loader.style.display = 'flex';
    }
}

function hideLoader() {
    const loader = document.getElementById('globalLoader');
    if(loader) {
        loader.style.display = 'none';
    }
}

// ==================== FIXED TRACK FUNCTION ====================
// This overrides any existing track function and ensures the modal shows
window.showApplicationTrack = async function(submissionId) {

    // Find the submission in our local data
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) {
        showNotification('Application not found', 'error');
        return;
    }

    // Create a simple tracking modal with HIGH z-index
    const modal = document.createElement('div');
    modal.className = 'modal active';
    modal.id = 'trackModal-' + Date.now();
    modal.style.cssText = 'display: flex !important; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999999; align-items: center; justify-content: center;';

    // Get stage info
    const stage = getApplicationStageSimple(sub);
    const stageColor = stage.includes('Reviewer') ? '#3b82f6' :
                      stage.includes('Approved') ? '#10b981' :
                      stage.includes('Rejected') ? '#ef4444' : '#f59e0b';
    const stageIcon = stage.includes('Reviewer') ? 'fa-user-check' :
                     stage.includes('Approved') ? 'fa-check-circle' :
                     stage.includes('Rejected') ? 'fa-times-circle' : 'fa-clock';

    modal.innerHTML = `
        <div class="modal-content" style="max-width: 500px; background: white; border-radius: 12px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.25); position: relative; z-index: 10000000;">
            <div class="modal-header" style="padding: 20px; border-bottom: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center;">
                <h2 style="margin:0; color:#003366;"><i class="fas fa-map-marker-alt"></i> Application Tracking</h2>
                <button class="close-btn" onclick="this.closest('.modal').remove()" style="background:none; border:none; font-size:24px; cursor:pointer;">&times;</button>
            </div>
            <div class="modal-body" style="padding: 30px;">
                <div style="text-align: center; margin-bottom: 30px;">
                    <i class="fas fa-file-alt" style="font-size: 60px; color: #003366;"></i>
                    <h3 style="margin: 15px 0 5px;">${sub.customerName || 'Unknown'}</h3>
                    <p style="color: #64748b;">Application ID: ${sub.id.substring(0, 8)}...</p>
                </div>
                <div style="background: #f8fafc; border-radius: 12px; padding: 20px;">
                    <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 15px;">
                        <div style="width: 40px; height: 40px; background: ${stageColor}; border-radius: 50%; display: flex; align-items: center; justify-content: center;">
                            <i class="fas ${stageIcon}" style="color: white;"></i>
                        </div>
                        <div>
                            <div style="font-size: 12px; color: #64748b;">Current Stage</div>
                            <div style="font-size: 18px; font-weight: 600; color: #1e293b;">${stage}</div>
                        </div>
                    </div>
                    <div style="border-top: 1px solid #e2e8f0; padding-top: 15px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                            <span style="color: #64748b;">Status:</span>
                            <span class="status-badge status-${sub.status}" style="padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; background: ${
                                sub.status === 'approved' ? '#d1fae5' :
                                sub.status === 'rejected' ? '#fee2e2' : '#fef3c7'
                            }; color: ${
                                sub.status === 'approved' ? '#065f46' :
                                sub.status === 'rejected' ? '#991b1b' : '#92400e'
                            };">${sub.status || 'pending'}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                            <span style="color: #64748b;">Uploaded:</span>
                            <span>${safeFormatDate(sub.uploadedAt)}</span>
                        </div>
                        ${sub.assignedTo ? `
                            <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                                <span style="color: #64748b;">Assigned To:</span>
                                <span>${sub.assignedTo}</span>
                            </div>
                        ` : ''}
                        ${sub.reviewedAt ? `
                            <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                                <span style="color: #64748b;">Reviewed:</span>
                                <span>${safeFormatDate(sub.reviewedAt)}</span>
                            </div>
                        ` : ''}
                        ${sub.comment ? `
                            <div style="margin-top: 15px;">
                                <span style="color: #64748b;">Comment:</span>
                                <p style="background: white; padding: 10px; border-radius: 8px; margin-top: 5px;">${sub.comment}</p>
                            </div>
                        ` : ''}
                    </div>
                </div>
                <div style="margin-top: 25px; text-align: center;">
                    <button class="action-btn" onclick="this.closest('.modal').remove()" style="padding: 10px 30px; background: #003366; color: white; border: none; border-radius: 6px; cursor: pointer;">Close</button>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(modal);

    // Close on backdrop click
    modal.addEventListener('click', (e) => {
        if (e.target === modal) modal.remove();
    });
};

// Simple stage determination function
function getApplicationStageSimple(submission) {
    if (!submission) return 'Unknown';
    if (submission.status === 'approved') return 'Approved - With RSA';
    if (submission.status === 'rejected') return 'Rejected - Fix Required';
    if (submission.assignedTo) return `With Reviewer: ${submission.assignedTo}`;
    return 'Pending Assignment';
}

// Make functions global
window.showLoader = showLoader;
window.hideLoader = hideLoader;
