import {
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { getSystemSettings } from './system-settings.js?v=20260507c';

export async function getMaintenanceSettings(db, { force = false } = {}) {
  try {
    const settings = await getSystemSettings(db, { force });
    return {
      maintenanceMode: settings.maintenanceMode === true,
      maintenanceMessage: String(settings.maintenanceMessage || '').trim(),
      updatedBy: '',
      updatedAt: null
    };
  } catch (_) {
    return {
      maintenanceMode: false,
      maintenanceMessage: 'System is currently under maintenance. Please try again later.',
      updatedBy: '',
      updatedAt: null
    };
  }
}

export function clearMaintenanceSettingsCache() {
  // Legacy no-op retained for compatibility.
}

export function isMaintenanceExemptRole(role = '') {
  return String(role || '').trim().toLowerCase() === 'super_admin';
}

function injectMaintenanceStyles() {
  if (document.getElementById('maintenanceModeStyles')) return;
  const style = document.createElement('style');
  style.id = 'maintenanceModeStyles';
  style.textContent = `
    .maintenance-mode-backdrop {
      position: fixed;
      inset: 0;
      z-index: 25000;
      background:
        radial-gradient(circle at top, rgba(30, 64, 175, 0.18), transparent 30%),
        rgba(2, 6, 23, 0.82);
      backdrop-filter: blur(12px);
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .maintenance-mode-card {
      width: min(620px, 100%);
      border-radius: 20px;
      background: linear-gradient(180deg, #ffffff, #f8fafc);
      box-shadow: 0 30px 80px rgba(15, 23, 42, 0.42);
      overflow: hidden;
      border: 1px solid rgba(226, 232, 240, 0.8);
    }
    .maintenance-mode-head {
      padding: 20px 24px;
      background: linear-gradient(135deg, #0f3b67, #0b2440);
      color: #fff;
      display: flex;
      align-items: center;
      gap: 14px;
    }
    .maintenance-mode-head i {
      font-size: 24px;
      color: #fbbf24;
    }
    .maintenance-mode-head strong {
      display: block;
      font-size: 21px;
      line-height: 1.2;
    }
    .maintenance-mode-head span {
      display: block;
      margin-top: 4px;
      font-size: 13px;
      opacity: 0.85;
    }
    .maintenance-mode-body {
      padding: 24px;
      color: #334155;
      line-height: 1.7;
    }
    .maintenance-mode-note {
      margin: 16px 0 0;
      padding: 14px 16px;
      border-left: 4px solid #f59e0b;
      border-radius: 12px;
      background: #fff7ed;
      color: #9a3412;
      font-size: 14px;
    }
    .maintenance-mode-actions {
      display: flex;
      justify-content: flex-end;
      gap: 12px;
      margin-top: 22px;
      flex-wrap: wrap;
    }
    .maintenance-mode-btn {
      border: 0;
      border-radius: 12px;
      padding: 11px 18px;
      font-weight: 700;
      cursor: pointer;
      font-size: 14px;
    }
    .maintenance-mode-btn.primary {
      background: #0f3b67;
      color: #fff;
    }
    .maintenance-mode-btn.secondary {
      background: #e2e8f0;
      color: #0f172a;
    }
  `;
  document.head.appendChild(style);
}

export function showMaintenanceOverlay(options = {}) {
  const existing = document.getElementById('maintenanceModeOverlay');
  if (existing) return existing;

  injectMaintenanceStyles();
  const overlay = document.createElement('div');
  overlay.id = 'maintenanceModeOverlay';
  overlay.className = 'maintenance-mode-backdrop';
  overlay.innerHTML = `
    <div class="maintenance-mode-card" role="dialog" aria-modal="true" aria-labelledby="maintenanceModeTitle">
      <div class="maintenance-mode-head">
        <i class="fas fa-screwdriver-wrench"></i>
        <div>
          <strong id="maintenanceModeTitle">System Under Maintenance</strong>
          <span>CMBank RSA portal is temporarily unavailable.</span>
        </div>
      </div>
      <div class="maintenance-mode-body">
        <p>${String(options.message || 'Maintenance mode is currently enabled. To protect live workflow data, portal access is temporarily restricted.').trim()}</p>
        <div class="maintenance-mode-note">Please try again later.</div>
        <div class="maintenance-mode-actions">
          <button type="button" class="maintenance-mode-btn secondary" id="maintenanceReloadBtn">Refresh</button>
          <button type="button" class="maintenance-mode-btn primary" id="maintenanceSignOutBtn">Sign Out</button>
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  document.body.style.overflow = 'hidden';

  overlay.querySelector('#maintenanceReloadBtn')?.addEventListener('click', () => {
    window.location.reload();
  });
  overlay.querySelector('#maintenanceSignOutBtn')?.addEventListener('click', async () => {
    if (typeof options.onSignOut === 'function') {
      await options.onSignOut();
      return;
    }
    window.location.href = 'index.html';
  });
  return overlay;
}
