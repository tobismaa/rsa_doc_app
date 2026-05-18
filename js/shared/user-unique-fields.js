import {
  doc,
  runTransaction,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

export function normalizeUniqueEmail(email) {
  return String(email || '').trim().toLowerCase();
}

export function normalizeUniqueWhatsapp(value) {
  return String(value || '').trim();
}

export function buildUniqueFieldKey(type, value) {
  return `${type}:${value}`;
}

function getUniqueFieldDoc(db, type, value) {
  return doc(db, 'userUniqueKeys', buildUniqueFieldKey(type, value));
}

function createDuplicateFieldError(field, value) {
  const error = new Error(`duplicate-${field}`);
  error.code = `duplicate-${field}`;
  error.field = field;
  error.value = value;
  return error;
}

export async function upsertUserWithUniqueFields(db, {
  userId,
  userData,
  previousUserData = null,
  merge = true
}) {
  const nextEmail = normalizeUniqueEmail(userData?.email);
  const nextWhatsapp = normalizeUniqueWhatsapp(userData?.whatsappNumber || userData?.phone);
  const previousEmail = normalizeUniqueEmail(previousUserData?.email);
  const previousWhatsapp = normalizeUniqueWhatsapp(previousUserData?.whatsappNumber || previousUserData?.phone);

  if (!userId) throw new Error('missing-user-id');
  if (!nextEmail) throw new Error('missing-email');

  const userRef = doc(db, 'users', userId);

  await runTransaction(db, async (transaction) => {
    const nextEmailRef = getUniqueFieldDoc(db, 'email', nextEmail);
    const nextEmailSnap = await transaction.get(nextEmailRef);

    if (nextEmailSnap.exists() && nextEmailSnap.data()?.userId !== userId) {
      throw createDuplicateFieldError('email', nextEmail);
    }

    if (nextWhatsapp) {
      const nextWhatsappRef = getUniqueFieldDoc(db, 'whatsapp', nextWhatsapp);
      const nextWhatsappSnap = await transaction.get(nextWhatsappRef);

      if (nextWhatsappSnap.exists() && nextWhatsappSnap.data()?.userId !== userId) {
        throw createDuplicateFieldError('whatsapp', nextWhatsapp);
      }

      transaction.set(nextWhatsappRef, {
        type: 'whatsapp',
        value: nextWhatsapp,
        userId,
        updatedAt: serverTimestamp()
      }, { merge: true });
    }

    transaction.set(userRef, userData, { merge });
    transaction.set(nextEmailRef, {
      type: 'email',
      value: nextEmail,
      userId,
      updatedAt: serverTimestamp()
    }, { merge: true });

    if (previousEmail && previousEmail !== nextEmail) {
      const previousEmailRef = getUniqueFieldDoc(db, 'email', previousEmail);
      const previousEmailSnap = await transaction.get(previousEmailRef);
      if (previousEmailSnap.exists() && previousEmailSnap.data()?.userId === userId) {
        transaction.delete(previousEmailRef);
      }
    }

    if (previousWhatsapp && previousWhatsapp !== nextWhatsapp) {
      const previousWhatsappRef = getUniqueFieldDoc(db, 'whatsapp', previousWhatsapp);
      const previousWhatsappSnap = await transaction.get(previousWhatsappRef);
      if (previousWhatsappSnap.exists() && previousWhatsappSnap.data()?.userId === userId) {
        transaction.delete(previousWhatsappRef);
      }
    }
  });
}

export async function deleteUserWithUniqueFields(db, {
  userId,
  userData = null
}) {
  const normalizedEmail = normalizeUniqueEmail(userData?.email);
  const normalizedWhatsapp = normalizeUniqueWhatsapp(userData?.whatsappNumber || userData?.phone);
  const userRef = doc(db, 'users', userId);

  await runTransaction(db, async (transaction) => {
    transaction.delete(userRef);

    if (normalizedEmail) {
      const emailRef = getUniqueFieldDoc(db, 'email', normalizedEmail);
      const emailSnap = await transaction.get(emailRef);
      if (emailSnap.exists() && emailSnap.data()?.userId === userId) {
        transaction.delete(emailRef);
      }
    }

    if (normalizedWhatsapp) {
      const whatsappRef = getUniqueFieldDoc(db, 'whatsapp', normalizedWhatsapp);
      const whatsappSnap = await transaction.get(whatsappRef);
      if (whatsappSnap.exists() && whatsappSnap.data()?.userId === userId) {
        transaction.delete(whatsappRef);
      }
    }
  });
}
