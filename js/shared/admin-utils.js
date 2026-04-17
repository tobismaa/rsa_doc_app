// Shared admin utility functions extracted from admin.js
export function normalizeUserRole(role) {
	const REVIEWER_ROLE_ALIASES = new Set(['reviewer', 'viewer']);
	const normalized = String(role || '').trim().toLowerCase();
	if (REVIEWER_ROLE_ALIASES.has(normalized)) return 'reviewer';
	return normalized || 'uploader';
}

export function getRoleLabel(role) {
	const normalized = normalizeUserRole(role);
	if (normalized === 'super_admin') return 'Super Admin';
	if (normalized === 'rsa') return 'RSA';
	if (normalized === 'payment') return 'Payment';
	return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

export function showNotification(msg, type) {
	// ...implementation as in admin.js...
}

export function showConfirmModal(title, msg, onConfirm) {
	// ...implementation as in admin.js...
}

export function closeConfirmModal() {
	// ...implementation as in admin.js...
}

export function runWithButtonSpinner(btnId, text, fn) {
	// ...implementation as in admin.js...
}

export function formatDate(date) {
	// ...implementation as in admin.js...
}

export function renderWhatsAppContactCell(user) {
	// ...implementation as in admin.js...
}

export function formatDepartment(dept) {
	// ...implementation as in admin.js...
}

export function escapeHtml(text) {
	// ...implementation as in admin.js...
}

export function getDisplayNameByEmail(email) {
	// ...implementation as in admin.js...
}

export function formatStatusLabel(status) {
	// ...implementation as in admin.js...
}

export function formatCurrency(value) {
	// ...implementation as in admin.js...
}

export function getSubmissionFinancials(sub) {
	// ...implementation as in admin.js...
}

export function showTestResultModal(title, results) {
	// ...implementation as in admin.js...
}
// Shared admin utility functions extracted from admin.js
