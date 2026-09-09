const statusCard = document.querySelector("#status-card");
const setupForm = document.querySelector("#setup-form");
const setupButton = document.querySelector("#setup-button");
const userForm = document.querySelector("#user-form");
const submitButton = document.querySelector("#submit-button");
const result = document.querySelector("#result");

function setStatus(type, message, icon) {
  statusCard.className = `status-card ${type}`;
  statusCard.innerHTML = `<i data-lucide="${icon}" aria-hidden="true"></i><span>${message}</span>`;
  window.lucide?.createIcons();
}

function setResult(type, message) {
  result.hidden = false;
  result.className = `result ${type}`;
  result.textContent = message;
}

function setBusy(isBusy) {
  submitButton.disabled = isBusy;
  submitButton.innerHTML = isBusy
    ? '<i data-lucide="loader-circle" aria-hidden="true"></i>Creating...'
    : '<i data-lucide="user-plus" aria-hidden="true"></i>Create User';
  window.lucide?.createIcons();
}

function setSetupBusy(isBusy) {
  setupButton.disabled = isBusy;
  setupButton.innerHTML = isBusy
    ? '<i data-lucide="loader-circle" aria-hidden="true"></i>Connecting...'
    : '<i data-lucide="key-round" aria-hidden="true"></i>Connect Admin Key';
  window.lucide?.createIcons();
}

async function checkStatus() {
  try {
    const response = await fetch("/api/status");
    const data = await response.json();

    if (!data.firebaseReady) {
      setupForm.hidden = false;
      userForm.hidden = true;
      setStatus("error", "Paste the Firebase private key JSON below, then this page will open the user form.", "circle-alert");
      return;
    }

    setupForm.hidden = true;
    userForm.hidden = false;
    setStatus("ready", `Connected to Firebase Admin${data.projectId ? ` for ${data.projectId}` : ""}.`, "shield-check");
  } catch (error) {
    setupForm.hidden = true;
    userForm.hidden = true;
    setStatus("error", error.message || "Could not reach the local admin server.", "circle-alert");
  }
}

setupForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  result.hidden = true;
  setSetupBusy(true);

  try {
    const response = await fetch("/api/service-account", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        serviceAccount: new FormData(setupForm).get("serviceAccount"),
      }),
    });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "Could not connect the admin key.");
    }

    setResult("success", "Admin key connected. You can create users now.");
    setupForm.reset();
    await checkStatus();
  } catch (error) {
    setResult("error", error.message || "Could not connect the admin key.");
  } finally {
    setSetupBusy(false);
  }
});

userForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  result.hidden = true;
  setBusy(true);

  const formData = new FormData(userForm);
  const payload = {
    email: formData.get("email"),
    password: formData.get("password"),
    displayName: formData.get("displayName"),
    role: formData.get("role"),
    active: formData.get("active") === "on",
    resetExistingPassword: formData.get("resetExistingPassword") === "on",
  };

  try {
    const response = await fetch("/api/users", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "Could not create user.");
    }

    const action = data.created ? "created" : "already existed and was approved";
    const passwordText = data.passwordUpdated ? " Password was reset." : "";
    const role = String(data.role || "staff").replace("_", " ");
    setResult("success", `${data.email} ${action} as ${role}. UID: ${data.uid}.${passwordText}`);
    userForm.reset();
    document.querySelector("#active").checked = true;
  } catch (error) {
    setResult("error", error.message || "Could not create user.");
  } finally {
    setBusy(false);
  }
});

window.addEventListener("DOMContentLoaded", () => {
  window.lucide?.createIcons();
  checkStatus();
});
