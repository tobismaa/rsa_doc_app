import { auth } from './firebase-config.js';
import {
    verifyPasswordResetCode,
    confirmPasswordReset
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";

const form = document.getElementById('resetPasswordForm');
const newPasswordInput = document.getElementById('newPassword');
const confirmInput = document.getElementById('confirmNewPassword');
const submitBtn = document.getElementById('resetSubmitBtn');
const messageEl = document.getElementById('resetMessage');
const strengthBar = document.getElementById('passwordStrengthBar');
const strengthText = document.getElementById('passwordStrengthText');

function showMessage(text, type = 'err') {
    if (!messageEl) return;
    messageEl.textContent = text;
    messageEl.className = `msg show ${type}`;
}

function getQueryParam(name) {
    return new URLSearchParams(window.location.search).get(name) || '';
}

function updateStrengthUi(password) {
    const v = String(password || '');
    let score = 0;
    if (v.length >= 6) score += 1;
    if (/[A-Z]/.test(v)) score += 1;
    if (/[0-9]/.test(v)) score += 1;
    if (/[^A-Za-z0-9]/.test(v)) score += 1;

    const pct = [0, 30, 55, 78, 100][score];
    if (strengthBar) strengthBar.style.width = `${pct}%`;

    if (!strengthText || !strengthBar) return;
    if (score <= 1) {
        strengthBar.style.background = '#ef4444';
        strengthText.textContent = 'Weak password';
    } else if (score === 2) {
        strengthBar.style.background = '#f59e0b';
        strengthText.textContent = 'Fair password';
    } else if (score === 3) {
        strengthBar.style.background = '#22c55e';
        strengthText.textContent = 'Good password';
    } else {
        strengthBar.style.background = '#16a34a';
        strengthText.textContent = 'Strong password';
    }
}

const oobCode = getQueryParam('oobCode');
const mode = getQueryParam('mode');

async function validateResetLink() {
    if (!oobCode || (mode && mode !== 'resetPassword')) {
        showMessage('This reset link is invalid or incomplete. Request a new one.', 'err');
        if (submitBtn) submitBtn.disabled = true;
        return false;
    }

    try {
        await verifyPasswordResetCode(auth, oobCode);
        return true;
    } catch (_) {
        showMessage('This reset link has expired or is invalid. Request a new one.', 'err');
        if (submitBtn) submitBtn.disabled = true;
        return false;
    }
}

if (form) {
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const newPassword = String(newPasswordInput?.value || '');
        const confirmPassword = String(confirmInput?.value || '');

        if (newPassword.length < 6) {
            showMessage('Password must be at least 6 characters.', 'err');
            return;
        }
        if (newPassword !== confirmPassword) {
            showMessage('Passwords do not match.', 'err');
            return;
        }

        if (submitBtn) submitBtn.disabled = true;
        try {
            await confirmPasswordReset(auth, oobCode, newPassword);
            showMessage('Password reset successful. You can now login.', 'ok');
            setTimeout(() => {
                window.location.href = 'index.html';
            }, 1800);
        } catch (error) {
            showMessage('Could not reset password: ' + (error?.message || 'Unknown error'), 'err');
        } finally {
            if (submitBtn) submitBtn.disabled = false;
        }
    });
}

if (newPasswordInput) {
    newPasswordInput.addEventListener('input', (e) => {
        updateStrengthUi(e.target?.value || '');
    });
}

document.querySelectorAll('.toggle-pass').forEach((btn) => {
    btn.addEventListener('click', () => {
        const targetId = btn.getAttribute('data-target');
        const input = targetId ? document.getElementById(targetId) : null;
        if (!input) return;
        const show = input.type === 'password';
        input.type = show ? 'text' : 'password';
        const icon = btn.querySelector('i');
        if (icon) {
            icon.classList.toggle('fa-eye', !show);
            icon.classList.toggle('fa-eye-slash', show);
        }
    });
});

validateResetLink();
