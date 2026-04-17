// User management logic extracted from admin.js
import { db, serverTimestamp, collection, query, where, orderBy, onSnapshot, getDocs, getDoc, doc, updateDoc, deleteDoc, addDoc } from './firebase-config.js';

// Utility imports (assume these are in shared/admin-utils.js or similar)
import { normalizeUserRole, getRoleLabel, showNotification, showConfirmModal, closeConfirmModal, runWithButtonSpinner, formatDate, renderWhatsAppContactCell, formatDepartment } from './shared/admin-utils.js';

let selectedUserId = null;
let allUsers = [];

function loadUsers() {
	const q = query(collection(db, 'users'), orderBy('createdAt', 'desc'));
	onSnapshot(q, (snapshot) => {
		const users = [];
		snapshot.forEach((doc) => {
			const userData = doc.data();
			const normalizedRole = normalizeUserRole(userData.role);
			if (userData.status !== 'pending' && normalizedRole !== 'super_admin') {
				const email = String(userData.email || '').toLowerCase();
				const displayName = userData.fullName || (email ? email.split('@')[0] : 'Unknown');
				users.push({
					id: doc.id,
					...userData,
					fullName: displayName
				});
			}
		});
		renderUsersTable(users);
		allUsers = users;
	}, (error) => {
		showNotification('Error loading users', 'error');
	});
}

function loadPendingUsers() {
	const q = query(collection(db, 'users'), where('status', '==', 'pending'));
	onSnapshot(q, (snapshot) => {
		const pendingUsers = [];
		snapshot.forEach((doc) => {
			const data = doc.data() || {};
			if (normalizeUserRole(data.role) === 'super_admin') return;
			pendingUsers.push({ id: doc.id, ...data });
		});
		renderPendingUsersGrid(pendingUsers);
		updatePendingUserCount(pendingUsers);
	}, (error) => {
		// fallback logic omitted for brevity
	});
}

function renderUsersTable(users) {
	// ...implementation as in admin.js...
}

function renderPendingUsersGrid(pendingUsers) {
	// ...implementation as in admin.js...
}

function updatePendingUserCount(users) {
	// ...implementation as in admin.js...
}

window.togglePendingApproval = (userId) => {
	// ...implementation as in admin.js...
};

window.activatePendingUser = (userId) => {
	// ...implementation as in admin.js...
};

window.rejectPendingUser = (userId) => {
	// ...implementation as in admin.js...
};

window.viewUser = async (userId) => {
	// ...implementation as in admin.js...
};

window.editUser = async (userId) => {
	// ...implementation as in admin.js...
};

// Export for use in admin.js if needed
export { loadUsers, loadPendingUsers };
// User management logic extracted from admin.js
