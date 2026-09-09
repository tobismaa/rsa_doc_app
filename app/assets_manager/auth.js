import { auth, db } from "./firebase-config.js";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";
import { doc, getDoc } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";

const allowedRoles = new Set(["staff", "admin", "super_admin"]);

export function friendlyError(error) {
  const messages = {
    "auth/invalid-credential": "The email or password is incorrect.",
    "auth/invalid-login-credentials": "The email or password is incorrect.",
    "auth/user-not-found": "The email or password is incorrect.",
    "auth/wrong-password": "The email or password is incorrect.",
    "auth/user-disabled": "This account is disabled. Please contact your administrator.",
    "auth/too-many-requests": "Too many attempts. Please wait before trying again.",
    "auth/network-request-failed": "Unable to connect. Check your internet connection and try again.",
    "auth/operation-not-allowed": "Email sign-in is not enabled. Please contact your administrator.",
    "auth/configuration-not-found": "Firebase Authentication is not turned on for this project yet. Open Firebase Authentication and click Get started once.",
    "permission-denied": "Your account does not have access to the asset register. Please contact your administrator.",
    "unavailable": "The asset service is unavailable. Please try again.",
    "assets/not-authorized": "Your account has not been granted Asset Manager access. Please contact your administrator.",
    "assets/duplicate-tag": "An asset with this tag already exists. Please use a unique tag.",
  };
  if (messages[error?.code]) return messages[error.code];
  if (String(error?.message || "").includes("configuration corresponding to the provided identifier")) {
    return messages["auth/configuration-not-found"];
  }
  return error?.message || "Unable to complete the request. Please try again or contact your administrator.";
}

async function waitForUser() {
  if (!auth) return null;
  await auth.authStateReady();
  return auth.currentUser;
}

async function checkAccess(user) {
  const profile = await getDoc(doc(db, "asset_manager_users", user.uid));
  const access = profile.exists() ? profile.data() : null;
  const role = access?.role ?? "staff";
  if (!access || access.active !== true || !allowedRoles.has(role)) {
    throw Object.assign(new Error("Asset Manager access required"), { code: "assets/not-authorized" });
  }
  return { ...access, role };
}

export async function requireAssetsSession() {
  if (!auth) {
    window.location.replace("login.html?error=config");
    return null;
  }
  const user = await waitForUser();
  if (!user) {
    window.location.replace("login.html");
    return null;
  }
  try {
    const access = await checkAccess(user);
    user.assetRole = access.role;
  } catch (error) {
    if (["assets/not-authorized", "permission-denied"].includes(error.code)) {
      await signOut(auth);
      window.location.replace("login.html?error=access");
      return null;
    }
    throw error;
  }
  onAuthStateChanged(auth, current => {
    if (!current) window.location.replace("login.html");
  });
  return user;
}

function showMessage(message) {
  const target = document.querySelector("#login-message");
  if (target) {
    target.textContent = message;
    target.classList.add("visible");
    target.classList.toggle("success", message.includes("successfully"));
  }
  const fixLink = document.querySelector("#firebase-fix-link");
  if (fixLink) {
    fixLink.classList.toggle("visible", message.includes("Firebase Authentication is not turned on"));
  }
  const returnLink = document.querySelector("#login-return-link");
  if (returnLink) {
    returnLink.classList.toggle("visible", message.includes("successfully"));
  }
}

async function initializeLogin() {
  const form = document.querySelector("#asset-login-form");
  if (!form) return;
  const password = document.querySelector("#password");
  const toggle = document.querySelector(".toggle-password");
  toggle?.addEventListener("click", () => {
    const show = password.type === "password";
    password.type = show ? "text" : "password";
    toggle.setAttribute("aria-label", show ? "Hide password" : "Show password");
    toggle.innerHTML = `<i data-lucide="${show ? "eye-off" : "eye"}" aria-hidden="true"></i>`;
    window.lucide?.createIcons();
  });
  const button = document.querySelector("#login-button");
  const signupButton = document.querySelector("#signup-button");
  if (!auth) {
    button.disabled = true;
    if (signupButton) signupButton.disabled = true;
    showMessage("Firebase setup is pending. Please contact your administrator.");
    return;
  }
  if (new URLSearchParams(location.search).get("error") === "access") {
    showMessage("Your account needs Asset Manager access. Please contact your administrator.");
  }
  if (new URLSearchParams(location.search).get("signup") === "success") {
    showMessage("Account created successfully. Please sign in.");
  }
  form.addEventListener("submit", async event => {
    event.preventDefault();
    if (button.disabled) return;
    button.disabled = true;
    button.querySelector("span").textContent = "Signing in...";
    try {
      const result = await signInWithEmailAndPassword(auth, document.querySelector("#email").value.trim(), password.value);
      const access = await checkAccess(result.user);
      result.user.assetRole = access.role;
      window.location.replace("index.html");
    } catch (error) {
      if (auth.currentUser) await signOut(auth).catch(() => {});
      showMessage(friendlyError(error));
    } finally {
      button.disabled = false;
      button.querySelector("span").textContent = "Sign in";
    }
  });
  signupButton?.addEventListener("click", async () => {
    const email = document.querySelector("#email").value.trim();
    const pass = password.value;

    if (!email || !pass) {
      showMessage("Enter email and password first, then click Sign up.");
      return;
    }

    button.disabled = true;
    signupButton.disabled = true;
    signupButton.querySelector("span").textContent = "Creating...";
    try {
      const response = await fetch("http://127.0.0.1:5055/api/signup", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          email,
          password: pass,
          displayName: email.split("@")[0],
          active: true,
          role: "staff",
          resetExistingPassword: false,
        }),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.ok) {
        throw Object.assign(new Error(data.error || "Could not create account."), { code: data.code });
      }

      if (auth.currentUser) await signOut(auth).catch(() => {});
      window.location.replace("login.html?signup=success");
    } catch (error) {
      if (auth.currentUser) await signOut(auth).catch(() => {});
      showMessage(friendlyError(error));
    } finally {
      button.disabled = false;
      signupButton.disabled = false;
      signupButton.querySelector("span").textContent = "Sign up";
    }
  });
  const user = await waitForUser();
  if (user) {
    try {
      const access = await checkAccess(user);
      user.assetRole = access.role;
      window.location.replace("index.html");
    } catch (error) {
      await signOut(auth).catch(() => {});
      showMessage(friendlyError(error));
    }
  }
}

const signOutButton = document.querySelector("#sign-out");
signOutButton?.addEventListener("click", async () => {
  signOutButton.disabled = true;
  try {
    if (auth) await signOut(auth);
    window.location.replace("login.html");
  } catch (error) {
    signOutButton.disabled = false;
    const message = document.querySelector("#app-message");
    message.textContent = friendlyError(error);
    message.hidden = false;
  }
});
initializeLogin().catch(error => showMessage(friendlyError(error)));
