// js/auth.js
import { auth, db } from './config.js';
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-auth.js";
import { doc, getDoc } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-firestore.js";

export async function loginUser(email, password) {
    try {
        const userCredential = await signInWithEmailAndPassword(auth, email, password);
        const user = userCredential.user;

        // Fetch User Role from Firestore 'users' collection
        const userDoc = await getDoc(doc(db, "users", user.uid));

        if (!userDoc.exists()) {
            await signOut(auth);
            throw new Error("User profile not found. Contact admin.");
        }

        const userData = userDoc.data();

        // Save session
        sessionStorage.setItem('userId', user.uid);
        sessionStorage.setItem('userRole', userData.role); // e.g., 'initiator', 'operations', 'coo'
        sessionStorage.setItem('userName', userData.name);

        // Redirect to Dashboard
        window.location.href = 'dashboard.html';
    } catch (error) {
        console.error("Login failed", error);
        throw error;
    }
}

export function logout() {
    sessionStorage.clear();
    signOut(auth);
    window.location.href = 'index.html';
}

// Protect Pages
export function requireAuth() {
    const user = sessionStorage.getItem('userId');
    if (!user) {
        window.location.href = 'index.html';
        return null;
    }
    return {
        id: user,
        role: sessionStorage.getItem('userRole'),
        name: sessionStorage.getItem('userName')
    };
}
