// Submission/document management logic extracted from admin.js
import { db, collection, query, orderBy, onSnapshot } from './firebase-config.js';
import { showNotification, formatDate, escapeHtml, getDisplayNameByEmail, formatCurrency, getSubmissionFinancials } from './shared/admin-utils.js';

let allSubmissions = [];

function loadSubmissions() {
	const q = query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc'));
	onSnapshot(q, (snapshot) => {
		allSubmissions = [];
		snapshot.forEach((doc) => {
			allSubmissions.push({ id: doc.id, ...doc.data() });
		});
		renderPendingDocs();
		renderApprovedDocs();
		renderRejectedDocs();
		renderTrackApplications();
		renderFinallySubmitted();
		renderPaymentQueue();
		updatePendingDocCount(allSubmissions.filter(s => s.status === 'pending'));
		updateRejectedDocCount(allSubmissions.filter(s => s.status === 'rejected'));
		updateFinallySubmittedCount();
		updatePaymentPendingCount();
	}, (error) => {
		// Silent fail
	});
}

function renderPendingDocs() {
	// ...implementation as in admin.js...
}
function renderApprovedDocs() {
	// ...implementation as in admin.js...
}
function renderRejectedDocs() {
	// ...implementation as in admin.js...
}
function renderTrackApplications() {
	// ...implementation as in admin.js...
}
function renderFinallySubmitted() {
	// ...implementation as in admin.js...
}
function renderPaymentQueue() {
	// ...implementation as in admin.js...
}
function updatePendingDocCount(items) {
	// ...implementation as in admin.js...
}
function updateRejectedDocCount(items) {
	// ...implementation as in admin.js...
}
function updateFinallySubmittedCount() {
	// ...implementation as in admin.js...
}
function updatePaymentPendingCount() {
	// ...implementation as in admin.js...
}

export { loadSubmissions };
// Submission/document management logic extracted from admin.js
