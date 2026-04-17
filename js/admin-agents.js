// Agent management logic extracted from admin.js
import { db, serverTimestamp, collection, query, where, orderBy, onSnapshot, getDocs, getDoc, doc, updateDoc, addDoc } from './firebase-config.js';
import { showNotification, showConfirmModal, closeConfirmModal, formatDate, escapeHtml, getDisplayNameByEmail, formatStatusLabel, runWithButtonSpinner, showTestResultModal } from './shared/admin-utils.js';

let allPendingAgents = [];
let allApprovedAgents = [];

function loadPendingAgents() {
	const q = query(collection(db, 'agents'), where('status', '==', 'pending'));
	onSnapshot(q, (snapshot) => {
		allPendingAgents = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
		renderPendingAgentsTable(allPendingAgents);
	}, () => {
		renderPendingAgentsTable([]);
	});
}

function loadApprovedAgents() {
	const q = query(collection(db, 'agents'), where('status', '==', 'approved'));
	onSnapshot(q, (snapshot) => {
		allApprovedAgents = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
		renderApprovedAgentsTable(allApprovedAgents);
	}, () => {
		renderApprovedAgentsTable([]);
	});
}

function renderPendingAgentsTable(items) {
	// ...implementation as in admin.js...
}

function renderApprovedAgentsTable(items) {
	// ...implementation as in admin.js...
}

window.viewPendingAgent = (agentId) => {
	// ...implementation as in admin.js...
};

window.viewApprovedAgent = (agentId) => {
	// ...implementation as in admin.js...
};

window.approveAgentRegistration = async (agentId, btnEl = null) => {
	// ...implementation as in admin.js...
};

window.rejectAgentRegistration = async (agentId, btnEl = null) => {
	// ...implementation as in admin.js...
};

export { loadPendingAgents, loadApprovedAgents };
// Agent management logic extracted from admin.js
