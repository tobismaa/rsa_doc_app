import { auth, db } from '../firebase-config.js?v=20260625c';
import { getCurrentUserProfile } from './user-directory.js?v=20260518a';

const params = new URL(window.location.href).searchParams;
const isSuperAdminView = String(params.get('view') || '').trim().toLowerCase() === 'super_admin';

async function shouldShowSwitcher(user) {
  if (!isSuperAdminView || !user) return false;
  try {
    const profile = await getCurrentUserProfile(db, user);
    return String(profile?.role || '').trim().toLowerCase() === 'super_admin';
  } catch (_) {
    return false;
  }
}

function renderSwitcher() {
  if (document.getElementById('superAdminDashboardSwitcher')) return;

  const dashboards = [
    { label: 'Super Admin', url: 'super-admin-dashboard.html' },
    { label: 'Uploader', url: 'dashboard.html?view=super_admin' },
    { label: 'Admin', url: 'admin-dashboard.html?view=super_admin' },
    { label: 'Reviewer', url: 'reviewer-dashboard.html?view=super_admin' },
    { label: 'RSA', url: 'rsa-dashboard.html?view=super_admin' },
    { label: 'Payment', url: 'payment-dashboard.html?view=super_admin' },
    { label: 'Audit', url: 'reports-monitoring-dashboard.html?view=super_admin' }
  ];

  const currentPath = window.location.pathname.split('/').pop() || 'dashboard.html';
  const wrap = document.createElement('div');
  wrap.id = 'superAdminDashboardSwitcher';
  wrap.setAttribute('aria-label', 'Super Admin dashboard switcher');
  wrap.style.cssText = [
    'position:fixed',
    'right:16px',
    'bottom:16px',
    'z-index:9999',
    'display:flex',
    'align-items:center',
    'gap:8px',
    'padding:10px 12px',
    'border:1px solid #cbd5e1',
    'border-radius:12px',
    'background:#ffffff',
    'box-shadow:0 12px 30px rgba(15,23,42,0.18)',
    'font:13px/1.4 Arial,sans-serif'
  ].join(';');

  const backButton = document.createElement('button');
  backButton.type = 'button';
  backButton.textContent = 'Back to Super Admin';
  backButton.style.cssText = [
    'padding:9px 12px',
    'border:0',
    'border-radius:8px',
    'background:#0f3b67',
    'color:#fff',
    'font-weight:700',
    'cursor:pointer'
  ].join(';');
  backButton.addEventListener('click', () => {
    window.location.href = 'super-admin-dashboard.html';
  });

  const select = document.createElement('select');
  select.style.cssText = 'padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:#0f172a;';
  dashboards.forEach((item) => {
    const option = document.createElement('option');
    option.value = item.url;
    option.textContent = item.label;
    const itemPath = item.url.split('?')[0];
    if (itemPath === currentPath) option.selected = true;
    select.appendChild(option);
  });
  select.addEventListener('change', () => {
    window.location.href = select.value;
  });

  wrap.appendChild(backButton);
  wrap.appendChild(select);

  const mount = () => {
    document.body.appendChild(wrap);
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mount, { once: true });
  } else {
    mount();
  }
}

auth.onAuthStateChanged(async (user) => {
  if (await shouldShowSwitcher(user)) {
    renderSwitcher();
  }
});
