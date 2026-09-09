// js/app.js
import { requireAuth, logout } from './auth.js';
import { db, WORKFLOW_STEPS } from './config.js';
import { collection, query, where, getDocs, orderBy, limit } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-firestore.js";

// Make logout available globally for the HTML button
window.logout = logout;

const user = requireAuth();
if (!user) return;

// 1. Setup User Info
document.getElementById('userInfo').textContent = `${user.name} (${user.role.toUpperCase()})`;

// 2. Build Navigation Based on Role
const nav = document.getElementById('navMenu');
const menuItems = [
    { label: 'Dashboard', url: 'dashboard.html', roles: ['all'] },
    { label: 'New Voucher', url: 'create-voucher.html', roles: ['initiator'] },
    { label: 'My Vouchers', url: 'dashboard.html', roles: ['all'] },
    { label: 'Approvals', url: 'dashboard.html', roles: ['iad', 'fincon', 'coo'] }, // Reuses dashboard but filters tasks
    { label: 'Posting', url: 'dashboard.html', roles: ['operations'] },
    { label: 'Settlement', url: 'dashboard.html', roles: ['initiator'] }
];

let navHTML = '';
menuItems.forEach(item => {
    if (item.roles.includes('all') || item.roles.includes(user.role)) {
        navHTML += `<a href="${item.url}" class="nav-link">${item.label}</a>`;
    }
});
nav.innerHTML = navHTML;

// 3. Load Dashboard Content Based on Role
loadDashboardContent();

async function loadDashboardContent() {
    const statsContainer = document.getElementById('statsContainer');
    const actionContent = document.getElementById('actionContent');
    const actionTitle = document.getElementById('actionTitle');
    const tableBody = document.getElementById('voucherTableBody');

    // --- A. Stats Logic ---
    let statsHTML = '';
    // Example: Count pending tasks for this user
    // (In real app, query Firestore here)
    statsHTML += `<div class="stat-box"><h3>Pending Tasks</h3><p>0</p></div>`;
    statsHTML += `<div class="stat-box"><h3>Total Processed</h3><p>0</p></div>`;
    statsContainer.innerHTML = statsHTML;

    // --- B. Action Logic (Role Specific) ---
    if (user.role === 'initiator') {
        actionTitle.textContent = "Quick Actions";
        actionContent.innerHTML = `<button class="btn" onclick="window.location.href='create-voucher.html'">+ Create Voucher</button>`;
    } else if (['iad', 'fincon', 'coo'].includes(user.role)) {
        actionTitle.textContent = "Pending Approvals";
        actionContent.innerHTML = `<p>You have pending approvals to review.</p>`;
    } else if (user.role === 'operations') {
        actionTitle.textContent = "Pending Posting";
        actionContent.innerHTML = `<p>Vouchers ready for reference number entry.</p>`;
    } else {
        actionTitle.textContent = "My Tasks";
        actionContent.innerHTML = `<p>No specific actions required.</p>`;
    }

    // --- C. Load Recent Vouchers (Unified Table) ---
    // Query vouchers relevant to this user
    const vouchersRef = collection(db, "vouchers");
    let q;

    if (user.role === 'initiator') {
        q = query(vouchersRef, where("initiatorId", "==", user.id), orderBy("createdAt", "desc"), limit(5));
    } else {
        // For others, show vouchers assigned to their role step
        // Note: This requires a field like 'currentStep' in voucher doc
        q = query(vouchersRef, where("currentStep", "==", user.role), orderBy("createdAt", "desc"), limit(5));
    }

    try {
        const snapshot = await getDocs(q);
        if (snapshot.empty) {
            tableBody.innerHTML = `<tr><td colspan="5">No vouchers found</td></tr>`;
        } else {
            let rows = '';
            snapshot.forEach(doc => {
                const data = doc.data();
                rows += `
                    <tr>
                        <td>${doc.id.substr(0,8)}...</td>
                        <td><span class="badge">${data.currentStep || 'initiated'}</span></td>
                        <td>${data.amount || 0}</td>
                        <td>${data.status || 'pending'}</td>
                        <td><a href="view-voucher.html?id=${doc.id}" style="color:#2563eb">View</a></td>
                    </tr>
                `;
            });
            tableBody.innerHTML = rows;
        }
    } catch (err) {
        console.error("Error loading vouchers", err);
        tableBody.innerHTML = `<tr><td colspan="5">Error loading data</td></tr>`;
    }
}
