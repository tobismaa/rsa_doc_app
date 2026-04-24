import {
  addDoc,
  collection,
  doc,
  getDoc,
  getDocs,
  runTransaction,
  serverTimestamp,
  setDoc,
  updateDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { isActiveUserWithRole, normalizeEmail } from "./user-directory.js?v=20260423b";

export function routingRuleDocId(uploaderEmail) {
  return encodeURIComponent(normalizeEmail(uploaderEmail));
}

export async function getUploaderRoutingRule(db, uploaderEmail) {
  const normalizedUploader = normalizeEmail(uploaderEmail);
  if (!normalizedUploader) return null;
  try {
    const snapshot = await getDoc(doc(db, 'uploaderRoutingRules', routingRuleDocId(normalizedUploader)));
    if (!snapshot.exists()) return null;
    const data = snapshot.data() || {};
    if (data.enabled === false) return null;
    return {
      uploaderEmail: normalizedUploader,
      reviewerEmail: normalizeEmail(data.reviewerEmail),
      rsaEmail: normalizeEmail(data.rsaEmail),
      paymentEmail: normalizeEmail(data.paymentEmail)
    };
  } catch (_) {
    return null;
  }
}

export async function getUsersByRoles(db, roles = []) {
  const snapshot = await getDocs(collection(db, 'users'));
  const allowed = new Set(roles.map((role) => String(role || '').toLowerCase()));
  return snapshot.docs
    .map((entry) => entry.data() || {})
    .filter((data) => {
      const role = String(data.role || '').toLowerCase();
      const status = String(data.status || 'active').toLowerCase();
      const leaveStatus = String(data.leaveStatus || '').toLowerCase();
      return allowed.has(role) && status !== 'deactivated' && leaveStatus !== 'on_leave';
    });
}

export async function getViewerEmails(db) {
  return (await getUsersByRoles(db, ['reviewer']))
    .map((data) => data.email)
    .filter(Boolean)
    .sort();
}

export async function assignRoundRobin({ db, currentUser, subRef, counterDoc }) {
  let uploaderEmail = normalizeEmail(currentUser?.email);
  if (!uploaderEmail) {
    try {
      const subSnap = await getDoc(subRef);
      uploaderEmail = normalizeEmail(subSnap.exists() ? subSnap.data()?.uploadedBy : '');
    } catch (_) {}
  }

  const routingRule = await getUploaderRoutingRule(db, uploaderEmail);
  const mappedReviewer = routingRule?.reviewerEmail || '';
  if (mappedReviewer && await isActiveUserWithRole(db, mappedReviewer, ['reviewer'])) {
    await updateDoc(subRef, {
      assignedTo: mappedReviewer,
      assignmentMode: 'uploader_routing',
      assignmentUploader: uploaderEmail || ''
    });
    try {
      const subSnap = await getDoc(subRef);
      if (subSnap.exists()) {
        const subData = subSnap.data();
        await addDoc(collection(db, 'roundRobinAssignments'), {
          submissionId: subRef.id,
          customerName: subData.customerName || 'N/A',
          assignedTo: mappedReviewer,
          assignedBy: currentUser?.email || 'System',
          assignedAt: serverTimestamp(),
          uploadedBy: subData.uploadedBy || uploaderEmail || 'N/A',
          assignmentMethod: 'uploader_routing'
        });
      }
    } catch (_) {}
    return mappedReviewer;
  }

  const viewers = await getViewerEmails(db);
  if (!viewers.length) return null;

  let assigned = null;
  try {
    await runTransaction(db, async (tx) => {
      let lastIndex = -1;
      let lastDate = '';
      const today = new Date().toISOString().slice(0, 10);
      const counterSnap = await tx.get(counterDoc);
      if (counterSnap.exists()) {
        const data = counterSnap.data();
        lastIndex = typeof data.lastIndex === 'number' ? data.lastIndex : -1;
        lastDate = data.lastDate || '';
      }
      if (lastDate !== today) lastIndex = -1;
      const newIndex = (lastIndex + 1) % viewers.length;
      assigned = viewers[newIndex];
      tx.set(counterDoc, { lastIndex: newIndex, lastDate: today }, { merge: true });
      tx.update(subRef, { assignedTo: assigned, assignmentMode: 'round_robin' });
    });
  } catch (_) {
    assigned = viewers[0] || null;
    if (assigned) {
      await updateDoc(subRef, { assignedTo: assigned, assignmentMode: 'round_robin_fallback' });
    }
  }

  if (assigned) {
    try {
      const subSnap = await getDoc(subRef);
      if (subSnap.exists()) {
        const subData = subSnap.data();
        await addDoc(collection(db, 'roundRobinAssignments'), {
          submissionId: subRef.id,
          customerName: subData.customerName || 'N/A',
          assignedTo: assigned,
          assignedBy: currentUser?.email || 'System',
          assignedAt: serverTimestamp(),
          uploadedBy: subData.uploadedBy || 'N/A',
          assignmentMethod: 'round_robin'
        });
      }
    } catch (_) {}
  }

  return assigned;
}
