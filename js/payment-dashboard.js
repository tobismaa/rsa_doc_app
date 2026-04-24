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

function roundDownToNearestThousand(value) {
    const num = Number(value || 0);
    return Math.max(0, Math.floor(num / 1000) * 1000);
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
    const computed25 = roundDownToNearestThousand(rsaBalance * 0.25);
    const stored25 = parseMoney(details.rsa25Percent || sub?.rsa25Percent || 0);
    const twentyFive = stored25 ? roundDownToNearestThousand(stored25) : computed25;
    const commission2 = twentyFive * 0.02;
    const pfa = String(details.pfa || sub?.pfa || '').trim() || '-';
    return { pfa, twentyFive, commission2 };
}

function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
}

function toSafeDomId(value) {
    return String(value || '').replace(/[^a-zA-Z0-9_-]+/g, '_');
}

function getAgentPaymentKey(sub) {
    const agentId = String(sub?.agentId || '').trim();
    if (agentId) return `agent:${agentId}`;
    const agentName = String(sub?.agentName || '').trim().toLowerCase();
    const uploaderEmail = normalizeEmail(sub?.uploadedBy);
    return `fallback:${uploaderEmail}::${agentName || 'no-agent'}`;
}

function buildAgentPaymentGroups(records = []) {
    const groups = new Map();
    records.forEach((sub) => {
        const key = getAgentPaymentKey(sub);
        const existing = groups.get(key) || {
            key,
            agentId: String(sub?.agentId || '').trim(),
            agentName: String(sub?.agentName || '').trim() || 'No Agent',
            uploaderEmail: normalizeEmail(sub?.uploadedBy),
            uploaderName: String(sub?.uploadedBy || '').trim() || '-',
            agentAccountNumber: String(sub?.agentAccountNumber || '').trim() || '-',
            agentAccountBank: String(sub?.agentAccountBank || '').trim() || '-',
            submissions: [],
            customerNames: new Set(),
            pfas: new Set(),
            total25: 0,
            totalCommission: 0,
            latestPaidAt: null,
            latestClearedAt: null,
            latestQueueAt: null
        };
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        existing.submissions.push(sub);
        existing.customerNames.add(String(sub?.customerName || 'Unknown'));
        if (pfa && pfa !== '-') existing.pfas.add(pfa);
        existing.total25 += twentyFive;
        existing.totalCommission += commission2;

        const paidAtMs = sub?.paidAt?.toMillis ? sub.paidAt.toMillis() : new Date(sub?.paidAt || 0).getTime();
        const clearedAtMs = sub?.clearedAt?.toMillis ? sub.clearedAt.toMillis() : new Date(sub?.clearedAt || 0).getTime();
        const queueAtMs = sub?.rsaSubmittedAt?.toMillis ? sub.rsaSubmittedAt.toMillis() : new Date(sub?.rsaSubmittedAt || sub?.updatedAt || 0).getTime();
        if (Number.isFinite(paidAtMs) && paidAtMs > 0 && (!existing.latestPaidAt || paidAtMs > existing.latestPaidAt)) existing.latestPaidAt = paidAtMs;
        if (Number.isFinite(clearedAtMs) && clearedAtMs > 0 && (!existing.latestClearedAt || clearedAtMs > existing.latestClearedAt)) existing.latestClearedAt = clearedAtMs;
        if (Number.isFinite(queueAtMs) && queueAtMs > 0 && (!existing.latestQueueAt || queueAtMs > existing.latestQueueAt)) existing.latestQueueAt = queueAtMs;

        groups.set(key, existing);
    });

    return Array.from(groups.values())
        .map((group) => ({
            ...group,
            customerCount: group.submissions.length,
            customerNames: Array.from(group.customerNames),
            pfas: Array.from(group.pfas)
        }))
        .sort((a, b) => (b.latestQueueAt || b.latestPaidAt || b.latestClearedAt || 0) - (a.latestQueueAt || a.latestPaidAt || a.latestClearedAt || 0));
}

function renderAgentBreakdownTable(group, mode = 'queue') {
    const rows = group.submissions.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const status = String(sub?.status || '').toLowerCase();
        const statusLabel = status === 'cleared' ? 'Cleared' : status === 'paid' ? 'Paid' : 'Sent to PFA';
        const dateLabel = mode === 'paid'
            ? formatDateValue(sub?.paidAt || sub?.updatedAt)
            : formatDateValue(sub?.rsaSubmittedAt || sub?.updatedAt);

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(pfa)}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td><span class="status-badge status-approved">${escapeHtml(statusLabel)}</span></td>
                <td>${escapeHtml(dateLabel)}</td>
                <td><button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');

    return `
        <div class="agent-breakdown-panel">
            <div class="agent-breakdown-meta">
                <strong>${escapeHtml(group.agentName)}</strong>
                <span>${group.customerCount} customer(s)</span>
                <span>${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</span>
            </div>
            <div class="table-container">
                <table class="customers-table agent-breakdown-table">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>PFA</th>
                            <th>25% Balance</th>
                            <th>2% Commission</th>
                            <th>Status</th>
                            <th>Date</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
        </div>
    `;
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
        payments: 'Agent Payment Queue',
        'paid-customers': 'Paid Agents',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Agent Payment Queue';
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
    const groupedQueue = buildAgentPaymentGroups(paymentQueue);

    if (paymentPendingCount) {
        paymentPendingCount.textContent = String(groupedQueue.length);
        paymentPendingCount.style.display = 'inline-block';
    }

    if (groupedQueue.length === 0) {
        paymentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No payment records available</td></tr>';
        return;
    }

    paymentsTableBody.innerHTML = groupedQueue.map((group) => {
        const breakdownId = `payment-breakdown-${toSafeDomId(group.key)}`;
        const pfaLabel = group.pfas.length === 1 ? group.pfas[0] : `${group.pfas.length} PFAs`;
        const queueDate = formatDateValue(group.latestQueueAt ? new Date(group.latestQueueAt) : null);
        const actionHtml = `<button class="action-btn" style="background:#16a34a;color:#fff;border:none;" onclick="window.markAgentPaid('${group.key}')"><i class="fas fa-check-circle"></i> Mark Agent Paid</button>`;

        return `
            <tr>
                <td>
                    <strong>${escapeHtml(group.agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</div>
                </td>
                <td>${escapeHtml(group.uploaderEmail || '-')}</td>
                <td>${group.customerCount}</td>
                <td>${formatCurrency(group.total25)}</td>
                <td>${formatCurrency(group.totalCommission)}</td>
                <td><span class="status-badge status-approved">Payment Pending</span><div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(pfaLabel)} • ${escapeHtml(queueDate)}</div></td>
                <td><button class="action-btn agent-breakdown-toggle" onclick="window.togglePaymentAgentBreakdown('${breakdownId}', this)"><i class="fas fa-chevron-down"></i> Breakdown</button> ${actionHtml}</td>
            </tr>
            <tr id="${breakdownId}" class="agent-breakdown-row" style="display:none;">
                <td colspan="7">${renderAgentBreakdownTable(group, 'queue')}</td>
            </tr>
        `;
    }).join('');
}

function renderPaidCustomers() {
    if (!paidCustomersTableBody) return;

    const paidCustomers = getPaymentRecords().filter((sub) => String(sub.status || '').toLowerCase() === 'paid');
    const groupedPaid = buildAgentPaymentGroups(paidCustomers);
    if (paidCustomerCount) {
        paidCustomerCount.textContent = String(groupedPaid.length);
        paidCustomerCount.style.display = 'inline-block';
    }

    if (groupedPaid.length === 0) {
        paidCustomersTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No paid agent batches yet</td></tr>';
        return;
    }

    paidCustomersTableBody.innerHTML = groupedPaid.map((group) => {
        const breakdownId = `paid-breakdown-${toSafeDomId(group.key)}`;
        const paidDate = formatDateValue(group.latestPaidAt ? new Date(group.latestPaidAt) : null);
        return `
            <tr>
                <td>
                    <strong>${escapeHtml(group.agentName)}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(group.agentAccountBank)} • ${escapeHtml(group.agentAccountNumber)}</div>
                </td>
                <td>${escapeHtml(group.uploaderEmail || '-')}</td>
                <td>${group.customerCount}</td>
                <td>${formatCurrency(group.total25)}</td>
                <td>${formatCurrency(group.totalCommission)}</td>
                <td><span class="status-badge status-approved">Paid</span><div style="font-size:12px;color:#64748b;margin-top:4px;">${escapeHtml(paidDate)}</div></td>
                <td><button class="action-btn agent-breakdown-toggle" onclick="window.togglePaymentAgentBreakdown('${breakdownId}', this)"><i class="fas fa-chevron-down"></i> Breakdown</button> <button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.clearPaidAgent('${group.key}')"><i class="fas fa-check-double"></i> Settle Agent</button></td>
            </tr>
            <tr id="${breakdownId}" class="agent-breakdown-row" style="display:none;">
                <td colspan="7">${renderAgentBreakdownTable(group, 'paid')}</td>
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

window.markAgentPaid = async (groupKey) => {
    const agentItems = getPaymentRecords().filter((sub) => {
        const currentStatus = String(sub.status || '').toLowerCase();
        return getAgentPaymentKey(sub) === groupKey && (currentStatus === 'sent_to_pfa' || currentStatus === 'rsa_submitted' || sub.finalSubmitted === true || sub.rsaSubmitted === true);
    });
    if (!agentItems.length) {
        showNotification('Agent payment batch not found', 'error');
        return;
    }
    const group = buildAgentPaymentGroups(agentItems)[0];
    const confirmed = confirm(`Mark ${group.agentName || 'this agent'} as PAID for ${group.customerCount} customer(s)?`);
    if (!confirmed) return;

    try {
        await Promise.all(agentItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentUser?.email || ''
        })));

        await addDoc(collection(db, 'audit'), {
            action: 'agent_commission_paid',
            agentKey: group.key,
            agentId: group.agentId || '',
            agentName: group.agentName || '',
            customerCount: group.customerCount,
            totalCommission: group.totalCommission,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        notifyStatusChangePush({
            currentUser,
            submissionId: agentItems[0].id,
            customerName: group.agentName || 'this agent',
            newStatus: 'paid',
            statusLabel: 'Paid',
            actionLabel: 'Agent Commission Marked Paid',
            message: `Commission for ${group.agentName || 'this agent'} was marked as paid for ${group.customerCount} customer(s).`
        }).catch(() => {});
        showNotification(`Marked ${group.agentName || 'agent'} as paid`, 'success');
    } catch (error) {
        showNotification('Failed to mark as paid', 'error');
    }
};

window.clearPaidAgent = async (groupKey) => {
    const paidItems = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid' && getAgentPaymentKey(s) === groupKey);
    if (paidItems.length === 0) {
        showNotification('No paid agent records to settle', 'info');
        return;
    }
    const group = buildAgentPaymentGroups(paidItems)[0];
    const confirmed = confirm(`Settle commission for ${group.agentName || 'this agent'} across ${group.customerCount} customer(s)?`);
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
            action: 'agent_commission_cleared',
            agentKey: group.key,
            agentId: group.agentId || '',
            agentName: group.agentName || '',
            count: paidItems.length,
            totalCommission: group.totalCommission,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        await Promise.all(
            paidItems.map((sub) => notifyStatusChangePush({
                currentUser,
                submissionId: sub.id,
                customerName: sub.customerName || '',
                newStatus: 'cleared',
                statusLabel: 'Cleared',
                actionLabel: 'Application Cleared',
                message: `Application for ${sub.customerName || 'this customer'} was cleared successfully.`
            }).catch(() => {}))
        );

        showNotification(`Settled ${group.agentName || 'agent'} commission`, 'success');
    } catch (error) {
        showNotification('Failed to clear paid records', 'error');
    }
};

window.clearPaidSubmissions = async () => {
    const paidItems = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid');
    if (!paidItems.length) {
        showNotification('No paid agent records to settle', 'info');
        return;
    }
    const groups = buildAgentPaymentGroups(paidItems);
    const confirmed = confirm(`Settle commission for ${groups.length} paid agent group(s)?`);
    if (!confirmed) return;
    try {
        for (const group of groups) {
            const groupItems = paidItems.filter((sub) => getAgentPaymentKey(sub) === group.key);
            await Promise.all(groupItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
                status: 'cleared',
                clearedAt: serverTimestamp(),
                clearedBy: currentUser?.email || ''
            })));
        }
        await addDoc(collection(db, 'audit'), {
            action: 'all_agent_commissions_cleared',
            count: paidItems.length,
            groupCount: groups.length,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification(`Settled ${groups.length} paid agent group(s)`, 'success');
    } catch (error) {
        showNotification('Failed to clear paid records', 'error');
    }
};

window.togglePaymentAgentBreakdown = (rowId, btn) => {
    const row = document.getElementById(rowId);
    if (!row) return;
    const isOpen = row.style.display !== 'none';
    row.style.display = isOpen ? 'none' : 'table-row';
    if (btn) {
        const icon = btn.querySelector('i');
        if (icon) icon.className = isOpen ? 'fas fa-chevron-down' : 'fas fa-chevron-up';
        const textNode = Array.from(btn.childNodes).find((node) => node.nodeType === Node.TEXT_NODE);
        if (textNode) textNode.textContent = isOpen ? ' Breakdown' : ' Hide';
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
