// js/auth.js - COMPLETE WORKING VERSION WITH DEPARTMENT FIELD
import { auth, db } from './firebase-config.js';
import { EMAIL_API_BASE_URL } from './email-api-config.js';
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
    query,
    where,
    getDocs,
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
            let errorMessage = 'Failed to send reset email.';
            switch (error.code) {
                case 'auth/invalid-email':
                    errorMessage = 'Invalid email format.';
                    break;
                case 'auth/user-not-found':
                    errorMessage = 'No account found with this email.';
                    break;
                case 'auth/unauthorized-continue-uri':
                case 'auth/invalid-continue-uri':
                    errorMessage = `Reset page URL is not authorized in Firebase. Add ${window.location.hostname} to Firebase Authorized Domains.`;
                    break;
                case 'auth/too-many-requests':
                    errorMessage = 'Too many attempts. Please try again later.';
                    break;
                default:
                    errorMessage = 'Failed to send reset email: ' + (error.message || 'Unknown error');
            }
            showError(errorMessage);
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
            const usersQuery = query(collection(db, 'users'), where('email', '==', email));
            const querySnapshot = await getDocs(usersQuery);

            if (!querySnapshot.empty) {
                const userDoc = querySnapshot.docs[0];
                const userData = userDoc.data();

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

                // Redirect based on role
                showSuccess('Login successful! Redirecting...');
                
                setTimeout(() => {
                    if (userData.role === 'super_admin') {
                        window.location.href = 'super-admin-dashboard.html';
                    } else if (userData.role === 'admin') {
                        window.location.href = 'admin-dashboard.html';
                    } else if (userData.role === 'reviewer' || userData.role === 'viewer') {
                        window.location.href = 'reviewer-dashboard.html';
                    } else if (userData.role === 'payment') {
                        window.location.href = 'payment-dashboard.html';
                    } else if (userData.role === 'rsa') {
                        window.location.href = 'rsa-dashboard.html';
                    } else {
                        window.location.href = 'dashboard.html';
                    }
                }, 1500);
            } else {
                // Security hardening: never auto-provision app roles from client login.
                showError('User record not found. Please contact admin.');
                await signOut(auth);
            }

        } catch (error) {
            
            let errorMessage = '';
            
            switch (error.code) {
                case 'auth/user-not-found':
                    errorMessage = 'No account found with this email.';
                    break;
                case 'auth/wrong-password':
                    errorMessage = 'Incorrect password.';
                    break;
                case 'auth/invalid-email':
                    errorMessage = 'Invalid email format.';
                    break;
                case 'auth/user-disabled':
                    errorMessage = 'This account has been disabled.';
                    break;
                case 'auth/too-many-requests':
                    errorMessage = 'Too many failed attempts. Please try again later.';
                    break;
                default:
                    errorMessage = 'Login failed: ' + error.message;
            }
            
            showError(errorMessage);
        }
    });
}

// ==================== SIGN OUT FUNCTION ====================
window.signOutUser = async () => {
    try {
        await signOut(auth);
        window.location.href = 'index.html';
    } catch (error) {
        showError('Error signing out');
    }
};

// Make functions globally available with security
Object.defineProperty(window, 'signOutUser', {
    value: window.signOutUser,
    writable: false,
    configurable: false
});
