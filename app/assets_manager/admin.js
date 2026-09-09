import { auth } from "./firebase-config.js";

const adminApi = "http://127.0.0.1:5055/api";

async function request(path, options = {}) {
  const token = await auth?.currentUser?.getIdToken();
  if (!token) throw Object.assign(new Error("Sign in again to use super admin controls."), { code: "assets/not-authorized" });
  const response = await fetch(`${adminApi}${path}`, {
    ...options,
    headers: {
      authorization: `Bearer ${token}`,
      "content-type": "application/json",
      ...(options.headers || {}),
    },
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok || !data.ok) throw new Error(data.error || "The admin action could not be completed.");
  return data;
}

export function listManagedUsers() {
  return request("/users");
}

export function createManagedUser(payload) {
  return request("/users", { method: "POST", body: JSON.stringify(payload) });
}

export function updateManagedUser(uid, payload) {
  return request(`/users/${encodeURIComponent(uid)}`, { method: "PATCH", body: JSON.stringify(payload) });
}
