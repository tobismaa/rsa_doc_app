import { createServer } from "node:http";
import { existsSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import admin from "firebase-admin";

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const publicDir = resolve(__dirname, "public");
const keyPath = process.env.GOOGLE_APPLICATION_CREDENTIALS
  ? resolve(process.env.GOOGLE_APPLICATION_CREDENTIALS)
  : resolve(__dirname, "serviceAccountKey.json");
const port = Number(process.env.PORT || 5055);

let firebaseReady = false;
let firebaseError = "";
let projectId = "";

function initializeFirebase() {
  firebaseReady = false;
  firebaseError = "";
  projectId = "";

  if (admin.apps.length) {
    firebaseReady = true;
    projectId = admin.app().options.credential.projectId || "";
    return;
  }

  if (!existsSync(keyPath)) {
    firebaseError = `Missing service account key at ${keyPath}`;
    return;
  }

  try {
    admin.initializeApp({
      credential: admin.credential.cert(keyPath),
    });
    firebaseReady = true;
    projectId = admin.app().options.credential.projectId || "";
  } catch (error) {
    firebaseError = error.message || "Firebase Admin could not start.";
  }
}

function validateServiceAccount(body) {
  const rawKey = String(body.serviceAccount || "")
    .trim()
    .replaceAll("\\_", "_")
    .replaceAll("\\@", "@")
    .replace(/\[(https?:\/\/[^\]]+)\]\((https?:\/\/[^)]+)\)/g, "$1");

  if (!rawKey) {
    return { error: "Paste the service account JSON first." };
  }

  let key;
  try {
    key = JSON.parse(rawKey);
  } catch {
    return { error: "That is not valid JSON. Paste the full downloaded key file." };
  }

  const requiredFields = ["project_id", "client_email", "private_key"];
  const missingField = requiredFields.find((field) => !key[field]);
  if (missingField) {
    return { error: `The key is missing ${missingField}. Paste the full Firebase private key JSON.` };
  }

  if (key.project_id !== "assetmanager-f2ac7") {
    return { error: "This key is not for assetmanager-f2ac7." };
  }

  return { key };
}

async function saveServiceAccount(body) {
  if (firebaseReady) {
    return {
      status: 200,
      payload: { ok: true, firebaseReady, projectId },
    };
  }

  const input = validateServiceAccount(body);
  if (input.error) {
    return { status: 400, payload: { ok: false, error: input.error } };
  }

  await writeFile(keyPath, `${JSON.stringify(input.key, null, 2)}\n`, {
    encoding: "utf8",
    mode: 0o600,
    flag: "wx",
  });
  initializeFirebase();

  if (!firebaseReady) {
    return {
      status: 500,
      payload: { ok: false, error: firebaseError || "Saved the key, but Firebase Admin could not start." },
    };
  }

  return {
    status: 201,
    payload: { ok: true, firebaseReady, projectId },
  };
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET,POST,PATCH,OPTIONS",
    "access-control-allow-headers": "content-type",
  });
  response.end(JSON.stringify(payload));
}

async function readJson(request) {
  const chunks = [];
  let total = 0;

  for await (const chunk of request) {
    total += chunk.length;
    if (total > 1024 * 1024) {
      throw new Error("Request is too large.");
    }
    chunks.push(chunk);
  }

  const text = Buffer.concat(chunks).toString("utf8");
  return text ? JSON.parse(text) : {};
}

function validateUserInput(body) {
  const roles = new Set(["staff", "admin", "super_admin"]);
  const email = String(body.email || "").trim();
  const password = String(body.password || "");
  const displayName = String(body.displayName || "").trim();
  const active = body.active !== false;
  const role = roles.has(body.role) ? body.role : "staff";
  const resetExistingPassword = body.resetExistingPassword === true;

  if (!email || !email.includes("@")) {
    return { error: "Enter a valid email address." };
  }

  if (!password || password.length < 6) {
    return { error: "Password must be at least 6 characters." };
  }

  if (displayName.length > 100) {
    return { error: "Full name must be 100 characters or less." };
  }

  return { email, password, displayName, active, role, resetExistingPassword };
}

async function createOrApproveUser(body) {
  if (!firebaseReady) {
    return {
      status: 503,
      payload: {
        ok: false,
        error: firebaseError || "Firebase Admin is not ready.",
      },
    };
  }

  const input = validateUserInput(body);
  if (input.error) {
    return { status: 400, payload: { ok: false, error: input.error } };
  }

  const userFields = {
    email: input.email,
    password: input.password,
    disabled: false,
  };

  if (input.displayName) {
    userFields.displayName = input.displayName;
  }

  let user;
  let created = false;
  let passwordUpdated = false;

  try {
    user = await admin.auth().createUser(userFields);
    created = true;
  } catch (error) {
    if (error.code === "auth/configuration-not-found") {
      return {
        status: 503,
        payload: {
          ok: false,
          code: error.code,
          error: "Firebase Authentication is not turned on for this project yet. Open Firebase Authentication and click Get started once.",
        },
      };
    }

    if (error.code !== "auth/email-already-exists") {
      return {
        status: 400,
        payload: { ok: false, code: error.code, error: error.message || "Could not create user." },
      };
    }

    user = await admin.auth().getUserByEmail(input.email);
    const update = { disabled: false };
    if (input.displayName) {
      update.displayName = input.displayName;
    }
    if (input.resetExistingPassword) {
      update.password = input.password;
      passwordUpdated = true;
    }
    user = await admin.auth().updateUser(user.uid, update);
  }

  const now = admin.firestore.FieldValue.serverTimestamp();
  await admin
    .firestore()
    .collection("asset_manager_users")
    .doc(user.uid)
    .set(
      {
        active: input.active,
        role: input.role,
        email: user.email,
        display_name: user.displayName || input.displayName || null,
        updated_at: now,
        ...(created ? { created_at: now } : {}),
      },
      { merge: true },
    );

  return {
    status: created ? 201 : 200,
    payload: {
      ok: true,
      created,
      passwordUpdated,
      uid: user.uid,
      email: user.email,
      active: input.active,
      role: input.role,
    },
  };
}

async function requireSuperAdmin(request) {
  if (!firebaseReady) {
    return { error: Object.assign(new Error(firebaseError || "Firebase Admin is not ready."), { status: 503 }) };
  }

  const header = request.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";
  if (!token) return { error: Object.assign(new Error("A signed-in super admin is required."), { status: 401 }) };

  try {
    const decoded = await admin.auth().verifyIdToken(token);
    const profile = await admin.firestore().collection("asset_manager_users").doc(decoded.uid).get();
    const access = profile.exists ? profile.data() : null;
    if (!access || access.active !== true || access.role !== "super_admin") {
      return { error: Object.assign(new Error("Super admin access is required."), { status: 403 }) };
    }
    return { decoded, access };
  } catch (error) {
    return { error: Object.assign(new Error("The admin session is invalid or expired."), { status: 401, cause: error }) };
  }
}

async function listManagedUsers() {
  const [authUsers, profileSnapshot] = await Promise.all([
    admin.auth().listUsers(1000),
    admin.firestore().collection("asset_manager_users").get(),
  ]);
  const profiles = new Map(profileSnapshot.docs.map(item => [item.id, item.data()]));
  return authUsers.users.map(user => {
    const profile = profiles.get(user.uid) || {};
    return {
      uid: user.uid,
      email: user.email || profile.email || "",
      display_name: user.displayName || profile.display_name || "",
      role: profile.role || "staff",
      active: profile.active === true,
      disabled: user.disabled,
      created_at: profile.created_at || null,
      last_sign_in_at: user.metadata.lastSignInTime || null,
    };
  }).sort((left, right) => left.email.localeCompare(right.email));
}

async function updateManagedUser(uid, body, actorUid) {
  if (!uid) return { status: 400, payload: { ok: false, error: "A user id is required." } };
  if (uid === actorUid && body.active === false) {
    return { status: 400, payload: { ok: false, error: "You cannot suspend your own super admin account." } };
  }

  const user = await admin.auth().getUser(uid);
  const profileRef = admin.firestore().collection("asset_manager_users").doc(uid);
  const profile = await profileRef.get();
  const current = profile.exists ? profile.data() : {};
  const update = {};
  const authUpdate = {};

  if (typeof body.active === "boolean") {
    update.active = body.active;
    authUpdate.disabled = !body.active;
  }
  if (["staff", "admin", "super_admin"].includes(body.role)) update.role = body.role;
  if (typeof body.displayName === "string" && body.displayName.trim().length <= 100) {
    update.display_name = body.displayName.trim();
    authUpdate.displayName = body.displayName.trim();
  }
  if (body.password) {
    if (String(body.password).length < 6) return { status: 400, payload: { ok: false, error: "Password must be at least 6 characters." } };
    authUpdate.password = String(body.password);
  }
  if (!Object.keys(update).length && !Object.keys(authUpdate).length) {
    return { status: 400, payload: { ok: false, error: "No user changes were supplied." } };
  }

  if (Object.keys(authUpdate).length) await admin.auth().updateUser(uid, authUpdate);
  await profileRef.set({
    ...current,
    email: user.email || current.email || null,
    ...update,
    updated_at: admin.firestore.FieldValue.serverTimestamp(),
  }, { merge: true });
  return { status: 200, payload: { ok: true, uid, ...update, passwordUpdated: Boolean(body.password) } };
}

const mimeTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".svg", "image/svg+xml"],
]);

async function serveStatic(request, response) {
  const url = new URL(request.url, `http://${request.headers.host}`);
  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = normalize(resolve(join(publicDir, requestedPath)));

  if (!filePath.startsWith(publicDir)) {
    sendJson(response, 403, { ok: false, error: "Forbidden." });
    return;
  }

  try {
    const data = await readFile(filePath);
    response.writeHead(200, {
      "content-type": mimeTypes.get(extname(filePath)) || "application/octet-stream",
      "cache-control": "no-store",
    });
    response.end(data);
  } catch {
    sendJson(response, 404, { ok: false, error: "Not found." });
  }
}

initializeFirebase();

createServer(async (request, response) => {
  try {
    if (request.method === "OPTIONS") {
      sendJson(response, 204, {});
      return;
    }

    if (request.method === "GET" && request.url?.startsWith("/api/status")) {
      sendJson(response, 200, {
        ok: true,
        firebaseReady,
        projectId,
        keyPath,
        error: firebaseReady ? "" : firebaseError,
      });
      return;
    }

    if (request.method === "GET" && request.url?.startsWith("/api/users")) {
      const access = await requireSuperAdmin(request);
      if (access.error) {
        sendJson(response, access.error.status, { ok: false, error: access.error.message });
        return;
      }
      sendJson(response, 200, { ok: true, users: await listManagedUsers() });
      return;
    }

    if (request.method === "POST" && request.url === "/api/users") {
      const access = await requireSuperAdmin(request);
      if (access.error) {
        sendJson(response, access.error.status, { ok: false, error: access.error.message });
        return;
      }
      const result = await createOrApproveUser(await readJson(request));
      sendJson(response, result.status, result.payload);
      return;
    }

    if (request.method === "POST" && request.url === "/api/signup") {
      const body = await readJson(request);
      const result = await createOrApproveUser({
        ...body,
        active: true,
        role: "staff",
        resetExistingPassword: false,
      });
      sendJson(response, result.status, result.payload);
      return;
    }

    if (request.method === "PATCH" && request.url?.startsWith("/api/users/")) {
      const access = await requireSuperAdmin(request);
      if (access.error) {
        sendJson(response, access.error.status, { ok: false, error: access.error.message });
        return;
      }
      const uid = decodeURIComponent(request.url.slice("/api/users/".length));
      const result = await updateManagedUser(uid, await readJson(request), access.decoded.uid);
      sendJson(response, result.status, result.payload);
      return;
    }

    if (request.method === "POST" && request.url === "/api/service-account") {
      const result = await saveServiceAccount(await readJson(request));
      sendJson(response, result.status, result.payload);
      return;
    }

    if (request.method === "GET") {
      await serveStatic(request, response);
      return;
    }

    sendJson(response, 405, { ok: false, error: "Method not allowed." });
  } catch (error) {
    sendJson(response, 500, {
      ok: false,
      error: error.message || "Something went wrong.",
    });
  }
}).listen(port, "127.0.0.1", () => {
  console.log(`Asset Manager User Admin: http://127.0.0.1:${port}`);
  if (!firebaseReady) {
    console.log(firebaseError);
  }
});
