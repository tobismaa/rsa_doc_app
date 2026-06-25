// js/auth.js - COMPLETE WORKING VERSION WITH DEPARTMENT FIELD
import { auth, db } from './firebase-config.js';
import { notifyAdminPushEvent } from './push-alerts.js';
import { EMAIL_API_BASE_URL } from './email-api-config.js';
import { performAppLogout } from './shared/logout.js?v=20260625b';
import { getCurrentUserProfile as getCurrentUserProfileShared } from './shared/user-directory.js?v=20260518a';
import { getMaintenanceSettings, isMaintenanceExemptRole } from './shared/maintenance-mode.js?v=20260507a';
import {
    createUserWithEmailAndPassword,
    signInWithEmailAndPassword,
    sendPasswordResetEmail,
    signOut
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    setDoc,
    doc,
    getDoc,
    query,
    where,
    getDocs,
    updateDoc,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

// ==================== SECURITY FUNCTIONS ====================
function sanitizeInput(input) {
    if (!input || typeof input !== 'string') return '';
    return input.replace(/[<>]/g, '');
}

function escapeHtml(unsafe) {
    if (!unsafe) return '';
    return unsafe
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// ==================== DOM ELEMENTS ====================
const loginTab = document.getElementById('loginTab');
const signupTab = document.getElementById('signupTab');
const loginForm = document.getElementById('loginForm');
const signupForm = document.getElementById('signupForm');
const loginFormElement = document.getElementById('loginFormElement');
const signupFormElement = document.getElementById('signupFormElement');
const forgotPasswordBtn = document.getElementById('forgotPasswordBtn');
const forgotPasswordModal = document.getElementById('forgotPasswordModal');
const forgotPasswordEmail = document.getElementById('forgotPasswordEmail');
const forgotPasswordSendBtn = document.getElementById('forgotPasswordSendBtn');
const forgotPasswordCancelBtn = document.getElementById('forgotPasswordCancelBtn');
const messageModal = document.getElementById('messageModal');
const modalSpinner = document.getElementById('modalSpinner');
const modalIcon = document.getElementById('modalIcon');
const modalErrorIcon = document.getElementById('modalErrorIcon');
const modalMessage = document.getElementById('modalMessage');
const modalCloseBtn = document.getElementById('modalCloseBtn');

// Error message elements
const nameError = document.getElementById('nameError');
const locationError = document.getElementById('locationError');
const whatsappError = document.getElementById('whatsappError');
const deptError = document.getElementById('deptError');
const emailError = document.getElementById('emailError');
const passwordError = document.getElementById('passwordError');

// ==================== TAB SWITCHING ====================
if (loginTab && signupTab) {
    loginTab.addEventListener('click', () => {
        loginTab.classList.add('active');
        signupTab.classList.remove('active');
        loginForm.classList.add('active');
        signupForm.classList.remove('active');
    });

    signupTab.addEventListener('click', () => {
        signupTab.classList.add('active');
        loginTab.classList.remove('active');
        signupForm.classList.add('active');
        loginForm.classList.remove('active');
    });
}

// ==================== MODAL FUNCTIONS ====================
function showSpinner(message) {
    if (!messageModal) return;
    modalSpinner.style.display = 'block';
    modalIcon.style.display = 'none';
    modalErrorIcon.style.display = 'none';
    modalCloseBtn.style.display = 'none';
    modalMessage.textContent = message || 'Processing...';
    messageModal.classList.add('active');
}

function showSuccess(message) {
    if (!messageModal) return;
    modalSpinner.style.display = 'none';
    modalIcon.style.display = 'block';
    modalErrorIcon.style.display = 'none';
    modalCloseBtn.style.display = 'inline-block';
    modalMessage.textContent = message || 'Success!';
    messageModal.classList.add('active');
}

function showError(message) {
    if (!messageModal) return;
    modalSpinner.style.display = 'none';
    modalIcon.style.display = 'none';
    modalErrorIcon.style.display = 'block';
    modalCloseBtn.style.display = 'inline-block';
    modalMessage.textContent = message || 'An error occurred';
    messageModal.classList.add('active');
}

function hideModal() {
    if (messageModal) {
        messageModal.classList.remove('active');
    }
}

async function checkIfAuthEmailExists(email) {
    const normalizedEmail = String(email || '').trim().toLowerCase();
    if (!normalizedEmail || !validateEmail(normalizedEmail)) return null;

    try {
        const apiKey = auth?.app?.options?.apiKey;
        if (!apiKey) return null;

        const response = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:createAuthUri?key=${encodeURIComponent(apiKey)}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                identifier: normalizedEmail,
                continueUri: window.location.origin
            })
        });

        const result = await response.json().catch(() => ({}));
        if (!response.ok) return null;
        if (typeof result.registered === 'boolean') return result.registered;
        if (Array.isArray(result.signinMethods)) return result.signinMethods.length > 0;
        return null;
    } catch (_) {
        return null;
    }
}

function getReadableLoginErrorMessage(error, emailExists = null) {
    switch (error?.code) {
        case 'auth/user-not-found':
            return 'No account found with this email.';
        case 'auth/wrong-password':
            return 'Incorrect password. Please try again.';
        case 'auth/invalid-email':
            return 'Invalid email format.';
        case 'auth/user-disabled':
            return 'This account has been disabled.';
        case 'auth/too-many-requests':
            return 'Too many failed attempts. Please try again later.';
        case 'auth/invalid-credential':
        case 'auth/invalid-login-credentials':
            if (emailExists === false) return 'No account found with this email.';
            if (emailExists === true) return 'Incorrect password. Please try again.';
            return 'Incorrect email or password. Please try again.';
        default:
            return 'Login failed: ' + (error?.message || 'Unknown error');
    }
}

function getReadableResetErrorMessage(error, emailExists = null) {
    switch (error?.code) {
        case 'auth/invalid-email':
            return 'Invalid email format.';
        case 'auth/user-not-found':
            return 'No account found with this email.';
        case 'auth/unauthorized-continue-uri':
        case 'auth/invalid-continue-uri':
            return `Reset page URL is not authorized in Firebase. Add ${window.location.hostname} to Firebase Authorized Domains.`;
        case 'auth/too-many-requests':
            return 'Too many attempts. Please try again later.';
        default:
            if (emailExists === false) return 'No account found with this email.';
            return 'Failed to send reset email: ' + (error?.message || 'Unknown error');
    }
}

async function findWritableUserRef(user) {
    if (!user?.uid && !user?.email) return null;

    const preferredUid = String(user?.uid || '').trim();
    const normalizedEmail = String(user?.email || '').trim().toLowerCase();

    if (preferredUid) {
        const uidRef = doc(db, 'users', preferredUid);
        const uidSnap = await getDoc(uidRef);
        if (uidSnap.exists()) return uidRef;
    }

    if (normalizedEmail) {
        const emailQuery = query(collection(db, 'users'), where('email', '==', normalizedEmail));
        const emailSnapshot = await getDocs(emailQuery);
        if (!emailSnapshot.empty) {
            return doc(db, 'users', emailSnapshot.docs[0].id);
        }
    }

    return preferredUid ? doc(db, 'users', preferredUid) : null;
}

function getWritableUserRefFromProfile(user, profile) {
    const profileDocId = String(profile?.__docId || '').trim();
    if (profileDocId) return doc(db, 'users', profileDocId);
    const preferredUid = String(user?.uid || '').trim();
    return preferredUid ? doc(db, 'users', preferredUid) : null;
}

function wait(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

async function resolveLoginUserProfile(user, { attempts = 3 } = {}) {
    if (!user) {
        return { profile: null, state: 'missing', error: null };
    }

    const preferredUid = String(user.uid || '').trim();
    const normalizedEmail = String(user.email || '').trim().toLowerCase();
    let lastError = null;

    for (let attempt = 0; attempt < attempts; attempt += 1) {
        let hadReadError = false;

        try {
            const sharedProfile = await getCurrentUserProfileShared(db, user);
            if (sharedProfile) {
                return { profile: sharedProfile, state: 'found', error: null };
            }
        } catch (error) {
            lastError = error;
            hadReadError = true;
        }

        if (preferredUid) {
            try {
                const uidRef = doc(db, 'users', preferredUid);
                const uidSnap = await getDoc(uidRef);
                if (uidSnap.exists()) {
                    return {
                        profile: { __docId: uidSnap.id, ...(uidSnap.data() || {}) },
                        state: 'found',
                        error: null
                    };
                }
            } catch (error) {
                lastError = error;
                hadReadError = true;
            }
        }

        if (normalizedEmail) {
            try {
                const emailQuery = query(collection(db, 'users'), where('email', '==', normalizedEmail));
                const emailSnapshot = await getDocs(emailQuery);
                if (!emailSnapshot.empty) {
                    const matchedDoc = emailSnapshot.docs.find((snap) => {
                        const data = snap.data() || {};
                        return String(data.uid || '').trim() === preferredUid || String(snap.id || '').trim() === preferredUid;
                    }) || emailSnapshot.docs[0];

                    return {
                        profile: { __docId: matchedDoc.id, ...(matchedDoc.data() || {}) },
                        state: 'found',
                        error: null
                    };
                }
            } catch (error) {
                lastError = error;
                hadReadError = true;
            }
        }

        if (!hadReadError) {
            return { profile: null, state: 'missing', error: null };
        }

        if (attempt < attempts - 1) {
            await wait(300 * (attempt + 1));
        }
    }

    return { profile: null, state: 'error', error: lastError };
}

// Close modal when clicking close button or outside
if (modalCloseBtn) {
    modalCloseBtn.addEventListener('click', hideModal);
}

if (messageModal) {
    messageModal.addEventListener('click', (e) => {
        if (e.target === messageModal) {
            hideModal();
        }
    });
}

// ==================== FORGOT PASSWORD ====================
function openForgotPasswordModal() {
    if (!forgotPasswordModal) return;
    const prefill = sanitizeInput(document.getElementById('loginEmail')?.value.trim().toLowerCase());
    if (forgotPasswordEmail) forgotPasswordEmail.value = prefill || '';
    forgotPasswordModal.classList.add('active');
}

function closeForgotPasswordModal() {
    if (!forgotPasswordModal) return;
    forgotPasswordModal.classList.remove('active');
}

if (forgotPasswordBtn) {
    forgotPasswordBtn.addEventListener('click', openForgotPasswordModal);
}

if (forgotPasswordCancelBtn) {
    forgotPasswordCancelBtn.addEventListener('click', closeForgotPasswordModal);
}

if (forgotPasswordModal) {
    forgotPasswordModal.addEventListener('click', (e) => {
        if (e.target === forgotPasswordModal) closeForgotPasswordModal();
    });
}

if (forgotPasswordSendBtn) {
    forgotPasswordSendBtn.addEventListener('click', async () => {
        const email = sanitizeInput(forgotPasswordEmail?.value.trim().toLowerCase() || '');
        if (!email) {
            showError('Please enter your email address.');
            return;
        }

        closeForgotPasswordModal();
        showSpinner('Sending password reset link...');
        try {
            const emailExists = await checkIfAuthEmailExists(email);
            if (emailExists === false) {
                showError('No account found with this email.');
                return;
            }

            const apiBaseUrl = String(window.__EMAIL_API_BASE_URL__ || EMAIL_API_BASE_URL || '').trim().replace(/\/+$/, '');
            const resetPageUrl = `${window.location.origin}/reset-password.html`;

            if (apiBaseUrl) {
                const response = await fetch(`${apiBaseUrl}/api/public/password-reset-request`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        email,
                        resetPageUrl
                    })
                });

                const result = await response.json().catch(() => ({}));
                if (!response.ok || !result.ok) {
                    throw new Error(result.error || `Reset email request failed (${response.status})`);
                }
            } else {
                const actionCodeSettings = {
                    url: resetPageUrl,
                    handleCodeInApp: true
                };
                await sendPasswordResetEmail(auth, email, actionCodeSettings);
            }

            showSuccess(`Password reset link sent to ${email}. Check inbox/spam.`);
        } catch (error) {
            const emailExists = await checkIfAuthEmailExists(email);
            showError(getReadableResetErrorMessage(error, emailExists));
        }
    });
}

// ==================== VALIDATION FUNCTIONS ====================
function validateName(name) {
    const pattern = /^[A-Za-z\s]+$/;
    return pattern.test(name) && name.length >= 2;
}

function validateLocation(location) {
    const value = String(location || '').trim();
    return value.length >= 2;
}

function validateWhatsapp(whatsapp) {
    const value = String(whatsapp || '').trim();
    const compact = value.replace(/[\s\-()]/g, '');
    return /^\+?\d{10,15}$/.test(compact);
}


function validateEmail(email) {
    const pattern = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
    return pattern.test(String(email || '').trim().toLowerCase());
}

function validatePassword(password) {
    return password.length >= 6;
}

function capitalizeName(name) {
    return name.split(' ').map(word => 
        word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
    ).join(' ');
}

function capitalizeWords(text) {
    return String(text || '')
        .split(' ')
        .map((word) => {
            if (!word) return word;
            return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
        })
        .join(' ');
}

function capitalizeFirst(word) {
    if (!word) return '';
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

// ==================== REAL-TIME VALIDATION ====================
document.addEventListener('DOMContentLoaded', () => {
    // Name validation
    const fullNameInput = document.getElementById('fullName');
    if (fullNameInput) {
        fullNameInput.addEventListener('input', function() {
            if (this.value.length > 0 && !validateName(this.value)) {
                nameError.textContent = 'Only letters and spaces allowed';
                nameError.classList.add('show');
                this.classList.add('input-error');
            } else {
                nameError.textContent = '';
                nameError.classList.remove('show');
                this.classList.remove('input-error');
            }
        });
        
        fullNameInput.addEventListener('blur', function() {
            if (validateName(this.value)) {
                this.value = capitalizeName(this.value);
            }
        });
    }

    const locationInput = document.getElementById('signupLocation');
    if (locationInput) {
        locationInput.addEventListener('input', function() {
            if (this.value.length > 0 && !validateLocation(this.value)) {
                locationError.textContent = 'Please enter a valid location';
                locationError.classList.add('show');
                this.classList.add('input-error');
            } else {
                locationError.textContent = '';
                locationError.classList.remove('show');
                this.classList.remove('input-error');
            }
        });
        locationInput.addEventListener('blur', function() {
            if (validateLocation(this.value)) this.value = capitalizeWords(this.value);
        });
    }

    const whatsappInput = document.getElementById('signupWhatsappNumber');
    if (whatsappInput) {
        whatsappInput.addEventListener('input', function() {
            this.value = String(this.value || '').replace(/\D/g, '').slice(0, 10);
            if (this.value.length > 0 && !/^\d{10}$/.test(this.value)) {
                whatsappError.textContent = 'Enter exactly 10 digits';
                whatsappError.classList.add('show');
                this.classList.add('input-error');
            } else {
                whatsappError.textContent = '';
                whatsappError.classList.remove('show');
                this.classList.remove('input-error');
            }
        });
    }


    // Email validation
    const emailInput = document.getElementById('signupEmail');
    if (emailInput) {
        emailInput.addEventListener('input', function() {
            this.value = this.value.toLowerCase();
            if (this.value.length > 0 && !validateEmail(this.value)) {
                emailError.textContent = 'Please enter a valid email address';
                emailError.classList.add('show');
                this.classList.add('input-error');
            } else {
                emailError.textContent = '';
                emailError.classList.remove('show');
                this.classList.remove('input-error');
            }
        });
    }

    // Password match validation
    const passwordInput = document.getElementById('signupPassword');
    const confirmInput = document.getElementById('confirmPassword');
    
    if (passwordInput && confirmInput) {
        const checkPasswordMatch = () => {
            if (confirmInput.value.length > 0 && passwordInput.value !== confirmInput.value) {
                passwordError.textContent = 'Passwords do not match';
                passwordError.classList.add('show');
                confirmInput.classList.add('input-error');
            } else {
                passwordError.textContent = '';
                passwordError.classList.remove('show');
                confirmInput.classList.remove('input-error');
            }
        };
        
        passwordInput.addEventListener('input', checkPasswordMatch);
        confirmInput.addEventListener('input', checkPasswordMatch);
    }
});

// ==================== SIGNUP FORM HANDLER ====================
if (signupFormElement) {
    signupFormElement.addEventListener('submit', async (e) => {
        e.preventDefault();

        // Get and sanitize form values
        const fullName = sanitizeInput(document.getElementById('fullName')?.value.trim());
        const location = sanitizeInput(document.getElementById('signupLocation')?.value.trim());
        const whatsappCode = sanitizeInput(document.getElementById('signupWhatsappCode')?.value.trim());
        const whatsappNumber = sanitizeInput(document.getElementById('signupWhatsappNumber')?.value.trim());
        // department no longer collected
        const department = '';
        const email = sanitizeInput(document.getElementById('signupEmail')?.value.trim().toLowerCase());
        const password = document.getElementById('signupPassword')?.value;
        const confirmPassword = document.getElementById('confirmPassword')?.value;

        // Validate all fields exist
        if (!fullName || !location || !whatsappNumber || !email || !password || !confirmPassword) {
            showError('Please fill in all fields');
            return;
        }

        // Validate name
        if (!validateName(fullName)) {
            showError('Please enter a valid name (letters and spaces only)');
            return;
        }

        if (!validateLocation(location)) {
            showError('Please enter a valid location');
            return;
        }

        if (!/^\d{10}$/.test(whatsappNumber)) {
            showError('Please enter exactly 10 digits for WhatsApp number');
            return;
        }


        // Validate email
        if (!validateEmail(email)) {
            showError('Please enter a valid email address');
            return;
        }

        // Validate password
        if (!validatePassword(password)) {
            showError('Password must be at least 6 characters');
            return;
        }

        // Validate passwords match
        if (password !== confirmPassword) {
            showError('Passwords do not match!');
            return;
        }

        // Capitalize fields
        const capitalizedName = capitalizeName(fullName);
        const normalizedLocation = capitalizeWords(location);
        const normalizedCode = whatsappCode.startsWith('+') ? whatsappCode : `+${whatsappCode}`;
        const normalizedWhatsapp = `${normalizedCode}${whatsappNumber}`;
        const capitalizedDept = '';
        // Show spinner
        showSpinner('Creating your account...');

        try {
            // Check if email already exists
            const emailQuery = query(collection(db, 'users'), where('email', '==', email));
            const existingEmailUsers = await getDocs(emailQuery);
            if (!existingEmailUsers.empty) {
                showError('An account with this email already exists.');
                return;
            }

            // Check if WhatsApp number already exists
            const whatsappQuery = query(collection(db, 'users'), where('whatsappNumber', '==', normalizedWhatsapp));
            const existingWhatsappUsers = await getDocs(whatsappQuery);
            if (!existingWhatsappUsers.empty) {
                showError('An account with this WhatsApp number already exists.');
                return;
            }

            // Legacy fallback: older records may store number in "phone"
            const legacyPhoneQuery = query(collection(db, 'users'), where('phone', '==', normalizedWhatsapp));
            const existingLegacyPhoneUsers = await getDocs(legacyPhoneQuery);
            if (!existingLegacyPhoneUsers.empty) {
                showError('An account with this WhatsApp number already exists.');
                return;
            }

            // Create user in Firebase Auth
            const userCredential = await createUserWithEmailAndPassword(auth, email, password);
            const user = userCredential.user;


            try {
                // Check if this is the first user
                const allUsersSnapshot = await getDocs(collection(db, 'users'));
                const isFirstUser = allUsersSnapshot.empty;

                // Save user data to Firestore with department
                const userData = {
                    uid: user.uid,
                    fullName: capitalizedName,
                    location: normalizedLocation,
                    whatsappCode: normalizedCode,
                    whatsappLocalNumber: whatsappNumber,
                    whatsappNumber: normalizedWhatsapp,
                    phone: normalizedWhatsapp,
                    // department removed
                    email: email,
                    role: isFirstUser ? 'admin' : 'uploader',
                    status: isFirstUser ? 'active' : 'pending',
                    createdAt: serverTimestamp(),
                    createdBy: 'self'
                };

                await setDoc(doc(db, 'users', user.uid), userData, { merge: true });

                await notifyAdminPushEvent({
                    currentUser: user,
                    eventType: 'new_user_registration',
                    title: isFirstUser ? 'New Admin Account Created' : 'New User Registration',
                    body: isFirstUser
                        ? `${capitalizedName} created the first admin account.`
                        : `${capitalizedName} registered and is pending approval.`,
                    clickUrl: '/admin-dashboard.html',
                    meta: {
                        userId: user.uid,
                        fullName: capitalizedName,
                        email,
                        role: userData.role,
                        status: userData.status
                    }
                });

                // Success message
                if (isFirstUser) {
                    showSuccess('First user created as ADMIN! You can now login.');
                } else {
                    showSuccess('Registration successful! Your account is pending approval.');
                }

                // Reset form
                signupFormElement.reset();

                // Switch to login tab after 2 seconds
                setTimeout(() => {
                    hideModal();
                    if (loginTab) loginTab.click();
                }, 2000);

            } catch (firestoreError) {
                
                // If Firestore save fails, delete the Auth user
                try {
                    await user.delete();
                } catch (deleteError) {
                }
                
                if (firestoreError.code === 'permission-denied') {
                    showError('Database permission error. Please contact admin.');
                } else {
                    showError('Failed to save user data: ' + firestoreError.message);
                }
            }

        } catch (authError) {
            
            // Handle specific Firebase errors
            let errorMessage = '';
            
            switch (authError.code) {
                case 'auth/email-already-in-use':
                    errorMessage = 'This email is already registered. Please login instead.';
                    break;
                case 'auth/invalid-email':
                    errorMessage = 'Invalid email format.';
                    break;
                case 'auth/weak-password':
                    errorMessage = 'Password should be at least 6 characters.';
                    break;
                case 'auth/network-request-failed':
                    errorMessage = 'Network error. Please check your internet connection.';
                    break;
                case 'auth/too-many-requests':
                    errorMessage = 'Too many attempts. Please try again later.';
                    break;
                default:
                    errorMessage = 'Registration failed: ' + authError.message;
            }
            
            showError(errorMessage);
        }
    });
}

// ==================== LOGIN FORM HANDLER ====================
if (loginFormElement) {
    loginFormElement.addEventListener('submit', async (e) => {
        e.preventDefault();

        const email = sanitizeInput(document.getElementById('loginEmail')?.value.trim().toLowerCase());
        const password = document.getElementById('loginPassword')?.value;

        if (!email || !password) {
            showError('Please enter email and password');
            return;
        }

        showSpinner('Signing in...');

        try {
            // Sign in with Firebase
            const userCredential = await signInWithEmailAndPassword(auth, email, password);
            const user = userCredential.user;

            // Check user status in Firestore
            const profileResolution = await resolveLoginUserProfile(user);
            const userData = profileResolution.profile;

            if (userData) {

                // Check status
                if (userData.status === 'pending') {
                    showError('Your account is pending approval. Please wait for admin activation.');
                    await signOut(auth);
                    return;
                } else if (userData.status === 'deactivated') {
                    showError('Your account has been deactivated. Please contact admin.');
                    await signOut(auth);
                    return;
                }

                const maintenanceSettings = await getMaintenanceSettings(db, { force: true });
                if (maintenanceSettings.maintenanceMode && !isMaintenanceExemptRole(userData.role)) {
                    showError(maintenanceSettings.maintenanceMessage || 'System is currently under maintenance. Please try again later.');
                    await signOut(auth);
                    return;
                }

                try {
                    const writableUserRef = getWritableUserRefFromProfile(user, userData) || await findWritableUserRef(user);
                    if (writableUserRef) {
                        await updateDoc(writableUserRef, {
                            lastLoginAt: serverTimestamp(),
                            isOnline: true,
                            lastSeenAt: serverTimestamp()
                        });
                    }
                } catch (_) {
                    // Fail open so an analytics timestamp issue does not block login.
                }

                // Keep the sign-in spinner visible and route straight to the dashboard.
                if (userData.role === 'super_admin') {
                    window.location.href = 'super-admin-dashboard.html';
                } else if (userData.role === 'admin') {
                    window.location.href = 'admin-dashboard.html';
                } else if (userData.role === 'reports_monitoring') {
                    window.location.href = 'reports-monitoring-dashboard.html';
                } else if (userData.role === 'reviewer') {
                    window.location.href = 'reviewer-dashboard.html';
                } else if (userData.role === 'payment') {
                    window.location.href = 'payment-dashboard.html';
                } else if (userData.role === 'rsa') {
                    window.location.href = 'rsa-dashboard.html';
                } else {
                    window.location.href = 'dashboard.html';
                }
            } else {
                // Security hardening: never auto-provision app roles from client login.
                if (profileResolution.state === 'missing') {
                    showError('Your app profile is missing. Please contact admin.');
                } else {
                    showError('We could not load your profile right now. Please try again.');
                }
                await signOut(auth);
            }

        } catch (error) {
            const emailExists = await checkIfAuthEmailExists(email);
            showError(getReadableLoginErrorMessage(error, emailExists));
        }
    });
}

// ==================== SIGN OUT FUNCTION ====================
window.signOutUser = async () => {
    await performAppLogout({
        auth,
        beforeSignOut: async () => {
            const user = auth.currentUser;
            if (user) {
                const writableUserRef = await findWritableUserRef(user);
                if (writableUserRef) {
                    await updateDoc(writableUserRef, {
                        isOnline: false,
                        lastSeenAt: serverTimestamp(),
                        lastLogoutAt: serverTimestamp()
                    }).catch(() => {});
                }
            }
        }
    });
};

// Make functions globally available with security
Object.defineProperty(window, 'signOutUser', {
    value: window.signOutUser,
    writable: false,
    configurable: false
});
