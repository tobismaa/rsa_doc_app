import { auth, db } from './firebase-config.js';
import { onAuthStateChanged } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js';
import { collection, query, where, getDocs, doc, updateDoc, serverTimestamp } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js';

const REQUIRED_FIELDS = [
  { key: 'fullName', label: 'Full Name' },
  { key: 'location', label: 'Location' },
  { key: 'whatsappCode', label: 'WhatsApp Country Code' },
  { key: 'whatsappLocalNumber', label: 'WhatsApp Number' }
];

const WHATSAPP_COUNTRY_CODES = [
  { code: '+234', label: 'Nigeria', flag: '🇳🇬' },
  { code: '+233', label: 'Ghana', flag: '🇬🇭' },
  { code: '+254', label: 'Kenya', flag: '🇰🇪' },
  { code: '+27', label: 'South Africa', flag: '🇿🇦' },
  { code: '+1', label: 'United States', flag: '🇺🇸' },
  { code: '+44', label: 'United Kingdom', flag: '🇬🇧' },
  { code: '+971', label: 'United Arab Emirates', flag: '🇦🇪' }
];

function normalizeText(value) {
  return String(value || '').trim();
}

function missingFields(userData = {}) {
  return REQUIRED_FIELDS.filter((field) => !normalizeText(userData[field.key]));
}

function buildModal(missing = [], existing = {}) {
  if (document.getElementById('profileGuardModal')) return;

  const style = document.createElement('style');
  style.id = 'profileGuardStyles';
  style.textContent = `
    .profile-guard-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(2, 6, 23, 0.74);
      z-index: 20000;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
    }
    .profile-guard-modal {
      width: min(560px, 100%);
      background: #fff;
      border-radius: 14px;
      box-shadow: 0 24px 60px rgba(2, 6, 23, 0.4);
      overflow: hidden;
    }
    .profile-guard-head {
      padding: 16px 18px;
      background: #0f3b67;
      color: #fff;
      font-weight: 700;
    }
    .profile-guard-body {
      padding: 16px 18px;
      color: #334155;
      font-size: 14px;
    }
    .profile-guard-list {
      margin: 0 0 12px;
      padding-left: 18px;
    }
    .profile-guard-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .profile-guard-grid label {
      display: block;
      font-size: 12px;
      font-weight: 700;
      color: #334155;
    }
    .profile-guard-grid input {
      width: 100%;
      margin-top: 6px;
      padding: 10px 11px;
      border: 1px solid #d1d5db;
      border-radius: 8px;
      font-size: 14px;
    }
    .profile-guard-grid select {
      width: 100%;
      margin-top: 6px;
      padding: 10px 11px;
      border: 1px solid #d1d5db;
      border-radius: 8px;
      font-size: 14px;
      background: #fff;
    }
    .profile-guard-grid .span-2 {
      grid-column: span 2;
    }
    .profile-guard-foot {
      padding: 12px 18px 16px;
      display: flex;
      justify-content: flex-end;
    }
    .profile-guard-btn {
      border: 0;
      border-radius: 8px;
      background: #0f766e;
      color: #fff;
      font-weight: 700;
      padding: 10px 16px;
      cursor: pointer;
    }
    .profile-guard-btn:disabled {
      opacity: 0.55;
      cursor: not-allowed;
    }
    .profile-guard-error {
      color: #dc2626;
      margin: 8px 0 0;
      font-size: 12px;
      min-height: 16px;
    }
    @media (max-width: 640px) {
      .profile-guard-grid { grid-template-columns: 1fr; }
      .profile-guard-grid .span-2 { grid-column: span 1; }
    }
  `;
  document.head.appendChild(style);

  const backdrop = document.createElement('div');
  backdrop.id = 'profileGuardModal';
  backdrop.className = 'profile-guard-backdrop';
  const missingList = missing.map((f) => `<li>${f.label}</li>`).join('');
  const selectedCode = normalizeText(existing.whatsappCode || '+234');
  const countryCodeOptions = WHATSAPP_COUNTRY_CODES.map((item) => `
    <option value="${item.code}" ${item.code === selectedCode ? 'selected' : ''}>${item.flag} ${item.code} ${item.label}</option>
  `).join('');
  backdrop.innerHTML = `
    <div class="profile-guard-modal" role="dialog" aria-modal="true" aria-labelledby="profileGuardTitle">
      <div id="profileGuardTitle" class="profile-guard-head">Complete Your Profile</div>
      <div class="profile-guard-body">
        <p>Your profile has missing details. Update now to continue.</p>
        <ul class="profile-guard-list">${missingList}</ul>
        <div class="profile-guard-grid">
          <label class="span-2">Full Name
            <input id="pgFullName" type="text" value="${normalizeText(existing.fullName)}" placeholder="Enter full name">
          </label>
          <label class="span-2">Location
            <input id="pgLocation" type="text" value="${normalizeText(existing.location)}" placeholder="Enter location">
          </label>
          <label>WhatsApp Country Code
            <select id="pgWhatsappCode">${countryCodeOptions}</select>
          </label>
          <label>WhatsApp Number
            <input id="pgWhatsappLocal" type="text" value="${normalizeText(existing.whatsappLocalNumber)}" placeholder="10 digits">
          </label>
        </div>
        <p id="profileGuardError" class="profile-guard-error"></p>
      </div>
      <div class="profile-guard-foot">
        <button id="profileGuardSaveBtn" class="profile-guard-btn" type="button">Save and Continue</button>
      </div>
    </div>
  `;
  document.body.appendChild(backdrop);
}

async function findUserDocByUidOrEmail(uid, email) {
  const byUid = query(collection(db, 'users'), where('uid', '==', uid));
  const uidSnap = await getDocs(byUid);
  if (!uidSnap.empty) return uidSnap.docs[0];

  const byEmail = query(collection(db, 'users'), where('email', '==', String(email || '').toLowerCase()));
  const emailSnap = await getDocs(byEmail);
  if (!emailSnap.empty) return emailSnap.docs[0];

  return null;
}

async function enforceProfileCompletion(user) {
  const userDoc = await findUserDocByUidOrEmail(user.uid, user.email);
  if (!userDoc) return;

  const userData = userDoc.data() || {};
  const missing = missingFields(userData);
  if (!missing.length) return;

  buildModal(missing, userData);
  const saveBtn = document.getElementById('profileGuardSaveBtn');
  const errorEl = document.getElementById('profileGuardError');
  const backdrop = document.getElementById('profileGuardModal');
  if (!saveBtn || !backdrop) return;

  document.body.style.overflow = 'hidden';

  const save = async () => {
    const fullName = normalizeText(document.getElementById('pgFullName')?.value);
    const location = normalizeText(document.getElementById('pgLocation')?.value);
    const whatsappCode = normalizeText(document.getElementById('pgWhatsappCode')?.value);
    const whatsappLocalNumber = normalizeText(document.getElementById('pgWhatsappLocal')?.value).replace(/\D/g, '');

    if (!fullName || !location || !whatsappCode || !whatsappLocalNumber) {
      if (errorEl) errorEl.textContent = 'All fields are required.';
      return;
    }
    if (!/^\+?\d{1,4}$/.test(whatsappCode)) {
      if (errorEl) errorEl.textContent = 'Country code is invalid.';
      return;
    }
    if (!/^\d{10}$/.test(whatsappLocalNumber)) {
      if (errorEl) errorEl.textContent = 'WhatsApp number must be exactly 10 digits.';
      return;
    }

    const normalizedCode = whatsappCode.startsWith('+') ? whatsappCode : `+${whatsappCode}`;
    const whatsappNumber = `${normalizedCode}${whatsappLocalNumber}`;

    saveBtn.disabled = true;
    saveBtn.textContent = 'Saving...';
    if (errorEl) errorEl.textContent = '';
    try {
      const waDupSnap = await getDocs(query(collection(db, 'users'), where('whatsappNumber', '==', whatsappNumber)));
      const waDuplicate = waDupSnap.docs.find((d) => d.id !== userDoc.id);
      if (waDuplicate) {
        if (errorEl) errorEl.textContent = 'This WhatsApp number already exists.';
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save and Continue';
        return;
      }

      const phoneDupSnap = await getDocs(query(collection(db, 'users'), where('phone', '==', whatsappNumber)));
      const phoneDuplicate = phoneDupSnap.docs.find((d) => d.id !== userDoc.id);
      if (phoneDuplicate) {
        if (errorEl) errorEl.textContent = 'This WhatsApp number already exists.';
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save and Continue';
        return;
      }

      await updateDoc(doc(db, 'users', userDoc.id), {
        fullName,
        location,
        whatsappCode: normalizedCode,
        whatsappLocalNumber,
        whatsappNumber,
        phone: whatsappNumber,
        updatedAt: serverTimestamp()
      });

      backdrop.remove();
      document.body.style.overflow = '';
      window.location.reload();
    } catch (err) {
      if (errorEl) errorEl.textContent = err?.message || 'Unable to save details.';
      saveBtn.disabled = false;
      saveBtn.textContent = 'Save and Continue';
    }
  };

  saveBtn.addEventListener('click', save);
}

onAuthStateChanged(auth, async (user) => {
  if (!user) return;
  try {
    await enforceProfileCompletion(user);
  } catch (_err) {
    // Fail open to avoid blocking legitimate users on transient issues.
  }
});
