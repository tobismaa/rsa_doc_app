// js/auth.js - COMPLETE WORKING VERSION WITH DEPARTMENT FIELD
import { auth, db } from './firebase-config.js';
import {
    createUserWithEmailAndPassword,
    signInWithEmailAndPassword,
    signOut
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    addDoc,
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
const messageModal = document.getElementById('messageModal');
const modalSpinner = document.getElementById('modalSpinner');
const modalIcon = document.getElementById('modalIcon');
const modalErrorIcon = document.getElementById('modalErrorIcon');
const modalMessage = document.getElementById('modalMessage');
const modalCloseBtn = document.getElementById('modalCloseBtn');

// Error message elements
const nameError = document.getElementById('nameError');
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

// Forgot password removed (per request). Admin can reset passwords from Admin dashboard.

// ==================== VALIDATION FUNCTIONS ====================
function validateName(name) {
    const pattern = /^[A-Za-z\s]+$/;
    return pattern.test(name) && name.length >= 2;
}


function validateEmail(email) {
    return email.endsWith('@cmbankng.com') && email.length > 10;
}

function validatePassword(password) {
    return password.length >= 6;
}

function capitalizeName(name) {
    return name.split(' ').map(word => 
        word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
    ).join(' ');
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


    // Email validation
    const emailInput = document.getElementById('signupEmail');
    if (emailInput) {
        emailInput.addEventListener('input', function() {
            this.value = this.value.toLowerCase();
            if (this.value.length > 0 && !validateEmail(this.value)) {
                emailError.textContent = 'Please use your @cmbankng.com email';
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
        // department no longer collected
        const department = '';
        const email = sanitizeInput(document.getElementById('signupEmail')?.value.trim().toLowerCase());
        const password = document.getElementById('signupPassword')?.value;
        const confirmPassword = document.getElementById('confirmPassword')?.value;

        // Validate all fields exist
        if (!fullName || !email || !password || !confirmPassword) {
            showError('Please fill in all fields');
            return;
        }

        // Validate name
        if (!validateName(fullName)) {
            showError('Please enter a valid name (letters and spaces only)');
            return;
        }


        // Validate email
        if (!validateEmail(email)) {
            showError('Please use your official @cmbankng.com email address');
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
        const capitalizedDept = '';
        // Show spinner
        showSpinner('Creating your account...');

        try {
            // Check if email already exists
            const usersQuery = query(collection(db, 'users'), where('email', '==', email));
            const existingUsers = await getDocs(usersQuery);
            
            if (!existingUsers.empty) {
                showError('An account with this email already exists.');
                return;
            }

            // Create user in Firebase Auth
            const userCredential = await createUserWithEmailAndPassword(auth, email, password);
            const user = userCredential.user;

            console.log('User created in Auth:', user.uid);

            try {
                // Check if this is the first user
                const allUsersSnapshot = await getDocs(collection(db, 'users'));
                const isFirstUser = allUsersSnapshot.empty;

                // Save user data to Firestore with department
                const userData = {
                    uid: user.uid,
                    fullName: capitalizedName,
                    // department removed
                    email: email,
                    role: isFirstUser ? 'admin' : 'uploader',
                    status: isFirstUser ? 'active' : 'pending',
                    createdAt: serverTimestamp(),
                    createdBy: 'self'
                };

                console.log('Saving to Firestore:', userData);

                const docRef = await addDoc(collection(db, 'users'), userData);
                console.log('User saved to Firestore with ID:', docRef.id);

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
                console.error('Firestore error:', firestoreError);
                
                // If Firestore save fails, delete the Auth user
                try {
                    await user.delete();
                    console.log('Auth user deleted due to Firestore error');
                } catch (deleteError) {
                    console.error('Could not delete Auth user:', deleteError);
                }
                
                if (firestoreError.code === 'permission-denied') {
                    showError('Database permission error. Please contact admin.');
                } else {
                    showError('Failed to save user data: ' + firestoreError.message);
                }
            }

        } catch (authError) {
            console.error('Auth error details:', authError);
            
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
                    if (userData.role === 'admin') {
                        window.location.href = 'admin-dashboard.html';
                    } else if (userData.role === 'reviewer' || userData.role === 'viewer') {
                        window.location.href = 'reviewer-dashboard.html';
                    } else if (userData.role === 'rsa') {
                        window.location.href = 'rsa-dashboard.html';
                    } else {
                        window.location.href = 'dashboard.html';
                    }
                }, 1500);
            } else {
                // User exists in Auth but not in Firestore - create record
                try {
                    await addDoc(collection(db, 'users'), {
                        uid: user.uid,
                        email: email,
                        fullName: email.split('@')[0],
                        // department removed for auto-generated record
                        role: 'uploader',
                        status: 'active',
                        createdAt: serverTimestamp()
                    });
                    showSuccess('Login successful! Redirecting...');
                    setTimeout(() => {
                        window.location.href = 'dashboard.html';
                    }, 1500);
                } catch (e) {
                    showError('User record not found. Please contact admin.');
                    await signOut(auth);
                }
            }

        } catch (error) {
            console.error('Login error:', error);
            
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
        console.error('Sign out error:', error);
        showError('Error signing out');
    }
};

// Make functions globally available with security
Object.defineProperty(window, 'signOutUser', {
    value: window.signOutUser,
    writable: false,
    configurable: false
});
