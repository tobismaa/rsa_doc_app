import { auth, db } from './firebase-config.js';
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    addDoc,
    query,
    where,
    getDocs,
    onSnapshot,
    updateDoc,
    doc,
    orderBy,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { notifyStatusChangePush } from './status-push.js';

let currentUser = null;
let currentUserData = null;
let allSubmissions = [];

const pageTitle = document.getElementById('pageTitle');
const paymentUserName = document.getElementById('paymentUserName');
const paymentPendingCount = document.getElementById('paymentPendingCount');
const paidCustomerCount = document.getElementById('paidCustomerCount');
const paymentsTableBody = document.getElementById('paymentsTableBody');
const paidCustomersTableBody = document.getElementById('paidCustomersTableBody');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
const notification = document.getElementById('notification');

function showNotification(message, type = 'info') {
    if (!notification) return;
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    setTimeout(() => { notification.style.display = 'none'; }, 3000);
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

function parseMoney(value) {
    const raw = String(value ?? '').replace(/[^0-9.\-]/g, '');
    const num = Number(raw);
    return Number.isFinite(num) ? num : 0;
}

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (e) {
        return `₦${num.toLocaleString()}`;
    }
}

function formatDateValue(value) {
    if (!value) return '-';
    try {
        const date = value.toDate ? value.toDate() : new Date(value);
        if (Number.isNaN(date.getTime())) return '-';
        return date.toLocaleString('en-NG');
    } catch (_) {
        return '-';
    }
}

function getSubmissionFinancials(sub) {
    const details = sub?.customerDetails || {};
    const rsaBalance = parseMoney(details.rsaBalance || sub?.rsaBalance || 0);
    const computed25 = Math.ceil((rsaBalance * 0.25) / 100) * 100;
    const twentyFive = parseMoney(details.rsa25Percent || sub?.rsa25Percent || computed25);
    const commission2 = twentyFive * 0.02;
    const pfa = String(details.pfa || sub?.pfa || '').trim() || '-';
    return { pfa, twentyFive, commission2 };
}

function renderProfile() {
    const fullName = currentUserData?.fullName || currentUser?.displayName || currentUser?.email || 'N/A';
    const registeredAt = currentUserData?.createdAt ? formatDateValue(currentUserData.createdAt) : '-';
    const email = currentUserData?.email || currentUser?.email || 'N/A';
    const whatsapp = currentUserData?.whatsappNumber || currentUserData?.phone || '-';
    const location = currentUserData?.location || '-';
    const role = String(currentUserData?.role || 'payment');
    const status = String(currentUserData?.status || 'active');

    if (paymentUserName) paymentUserName.textContent = fullName;
    if (profileNameEl) profileNameEl.textContent = fullName;
    if (profileRegisteredAtEl) profileRegisteredAtEl.textContent = registeredAt;
    if (profileEmailEl) profileEmailEl.textContent = email;
    if (profileWhatsappEl) profileWhatsappEl.textContent = whatsapp;
    if (profileLocationEl) profileLocationEl.textContent = location;
    if (profileRoleEl) profileRoleEl.textContent = role.charAt(0).toUpperCase() + role.slice(1);
    if (profileStatusEl) profileStatusEl.textContent = status.charAt(0).toUpperCase() + status.slice(1);
}

function switchTab(tabId) {
    document.querySelectorAll('.nav-item').forEach((nav) => nav.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');
    const titles = {
        payments: 'Payment Queue',
        'paid-customers': 'Paid Customers',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Payment Queue';
}

function getPaymentRecords() {
    return allSubmissions.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        if (status === 'cleared') return false;
        if (status === 'sent_to_pfa' || status === 'rsa_submitted' || status === 'paid') return true;
        // Backward compatibility for legacy final-submitted records
        return sub.finalSubmitted === true || sub.rsaSubmitted === true;
    });
}

function renderPaymentQueue() {
    if (!paymentsTableBody) return;

    const paymentQueue = getPaymentRecords().filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        if (status === 'paid') return false;
        // Any submission sent to PFA should appear immediately in payment queue.
        return status === 'sent_to_pfa' || status === 'rsa_submitted' || sub.finalSubmitted === true || sub.rsaSubmitted === true;
    });

    const pendingCount = paymentQueue.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return status === 'sent_to_pfa' || status === 'rsa_submitted';
    }).length;
    if (paymentPendingCount) {
        paymentPendingCount.textContent = String(pendingCount);
        paymentPendingCount.style.display = pendingCount > 0 ? 'inline' : 'none';
    }

    if (paymentQueue.length === 0) {
        paymentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No payment records available</td></tr>';
        return;
    }

    paymentsTableBody.innerHTML = paymentQueue.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const status = String(sub.status || '').toLowerCase();
        const isPaid = status === 'paid';
        const statusLabel = isPaid ? 'Paid' : 'Sent to PFA';
        const actionHtml = isPaid
            ? '<button class="action-btn" style="opacity:.65;cursor:not-allowed;" disabled><i class="fas fa-check"></i> Paid</button>'
            : `<button class="action-btn" style="background:#16a34a;color:#fff;border:none;" onclick="window.markSubmissionPaid('${sub.id}')"><i class="fas fa-check-circle"></i> Paid</button>`;

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(pfa)}</td>
                <td>${escapeHtml(sub.agentName || '-')}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td><span class="status-badge status-approved">${statusLabel}</span></td>
                <td>${actionHtml} <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');
}

function renderPaidCustomers() {
    if (!paidCustomersTableBody) return;

    const paidCustomers = getPaymentRecords().filter((sub) => String(sub.status || '').toLowerCase() === 'paid');
    if (paidCustomerCount) {
        paidCustomerCount.textContent = String(paidCustomers.length);
        paidCustomerCount.style.display = paidCustomers.length > 0 ? 'inline' : 'none';
    }

    if (paidCustomers.length === 0) {
        paidCustomersTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No paid customers yet</td></tr>';
        return;
    }

    paidCustomersTableBody.innerHTML = paidCustomers.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(pfa)}</td>
                <td>${escapeHtml(sub.agentName || '-')}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td><span class="status-badge status-approved">Paid</span></td>
                <td><button class="action-btn" style="opacity:.65;cursor:not-allowed;" disabled><i class="fas fa-check"></i> Paid</button> <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');
}

function loadSubmissions() {
    const q = query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc'));
    onSnapshot(q, (snapshot) => {
        allSubmissions = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
        renderPaymentQueue();
        renderPaidCustomers();
    }, () => {
        showNotification('Failed to load payment queue', 'error');
    });
}

window.markSubmissionPaid = async (submissionId) => {
    const sub = allSubmissions.find((s) => s.id === submissionId);
    if (!sub) {
        showNotification('Submission not found', 'error');
        return;
    }
    const currentStatus = String(sub.status || '').toLowerCase();
    const allowed = currentStatus === 'sent_to_pfa' || currentStatus === 'rsa_submitted';
    if (!allowed) {
        showNotification('Only applications already sent to PFA can be marked as paid.', 'warning');
        return;
    }
    const confirmed = confirm(`Mark ${sub.customerName || 'this customer'} as PAID?`);
    if (!confirmed) return;

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentUser?.email || ''
        });

        await addDoc(collection(db, 'audit'), {
            action: 'submission_paid',
            submissionId,
            customerName: sub.customerName || '',
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        notifyStatusChangePush({
            currentUser,
            submissionId,
            customerName: sub.customerName || '',
            newStatus: 'paid',
            statusLabel: 'Paid'
        }).catch(() => {});
        showNotification('Marked as paid successfully', 'success');
    } catch (error) {
        showNotification('Failed to mark as paid', 'error');
    }
};

window.clearPaidSubmissions = async () => {
    const paidItems = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid');
    if (paidItems.length === 0) {
        showNotification('No paid records to clear', 'info');
        return;
    }
    const confirmed = confirm(`Clear ${paidItems.length} paid record(s)?`);
    if (!confirmed) return;

    try {
        await Promise.all(
            paidItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
                status: 'cleared',
                clearedAt: serverTimestamp(),
                clearedBy: currentUser?.email || ''
            }))
        );

        await addDoc(collection(db, 'audit'), {
            action: 'paid_records_cleared',
            count: paidItems.length,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification(`Cleared ${paidItems.length} paid record(s)`, 'success');
    } catch (error) {
        showNotification('Failed to clear paid records', 'error');
    }
};

window.signOutUser = async () => {
    try { await signOut(auth); } catch (e) {}
    window.location.href = 'index.html';
};

function forceHardRefresh() {
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
}

document.getElementById('forceRefreshBtn')?.addEventListener('click', (e) => {
    e.preventDefault();
    forceHardRefresh();
});

document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        switchTab(item.dataset.tab);
    });
});

auth.onAuthStateChanged(async (user) => {
    if (!user) {
        window.location.href = 'index.html';
        return;
    }
    currentUser = user;

    try {
        let userData = null;
        const uidQuery = query(collection(db, 'users'), where('uid', '==', user.uid));
        const uidSnap = await getDocs(uidQuery);
        if (!uidSnap.empty) {
            userData = uidSnap.docs[0].data();
        } else if (user.email) {
            const emailQuery = query(collection(db, 'users'), where('email', '==', user.email.toLowerCase()));
            const emailSnap = await getDocs(emailQuery);
            if (!emailSnap.empty) userData = emailSnap.docs[0].data();
        }

        if (!userData) {
            showNotification('User profile not found', 'error');
            window.location.href = 'index.html';
            return;
        }

        const role = String(userData.role || '').toLowerCase();
        if (role === 'admin') {
            window.location.href = 'admin-dashboard.html';
            return;
        }
        if (role !== 'payment') {
            window.location.href = 'index.html';
            return;
        }

        currentUserData = userData;
        renderProfile();
        loadSubmissions();
    } catch (error) {
        showNotification('Could not validate session', 'error');
        window.location.href = 'index.html';
    }
});
