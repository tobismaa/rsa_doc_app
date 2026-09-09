import { db, auth } from "./firebase-config.js";
import { collection, doc, getDocs, orderBy, query, runTransaction, serverTimestamp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";

export function validateAsset(payload) {
  if (!payload.name?.trim() || !payload.tag?.trim()) throw new Error("Enter an asset name and tag.");
  if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,99}$/.test(payload.tag)) {
    throw new Error("Use up to 100 letters, numbers, hyphens or underscores for the asset tag.");
  }
  if (payload.name.length > 200 || !payload.category || payload.category.length > 100) throw new Error("Check the asset name and category.");
  if (![payload.purchase_cost, payload.book_value].every(value => Number.isFinite(value) && value >= 0)) {
    throw new Error("Purchase cost and book value must be valid, non-negative amounts.");
  }
  if (!["Active", "Service due", "Review", "Disposal"].includes(payload.status)) throw new Error("Choose a valid asset status.");
  for (const field of ["custodian", "location"]) {
    if (payload[field] !== null && (typeof payload[field] !== "string" || payload[field].length > 200)) throw new Error(`Keep ${field} under 200 characters.`);
  }
  if (payload.purchase_date !== null && !/^\d{4}-\d{2}-\d{2}$/.test(payload.purchase_date)) throw new Error("Enter a valid purchase date.");
}

export async function listAssets() {
  const snapshot = await getDocs(query(collection(db, "assets"), orderBy("created_at", "desc")));
  return snapshot.docs.map(item => ({ ...item.data(), id: item.id }));
}

export async function createAsset(payload) {
  validateAsset(payload);
  if (!auth?.currentUser) throw Object.assign(new Error("Sign in required"), { code: "assets/not-authorized" });
  // The tag is the document ID: a transaction prevents concurrent duplicate tags.
  const reference = doc(db, "assets", payload.tag);
  await runTransaction(db, async transaction => {
    const existing = await transaction.get(reference);
    if (existing.exists()) throw Object.assign(new Error("Duplicate tag"), { code: "assets/duplicate-tag" });
    transaction.set(reference, {
      ...payload,
      created_by: auth.currentUser.uid,
      created_at: serverTimestamp(),
      updated_at: serverTimestamp(),
    });
  });
}
