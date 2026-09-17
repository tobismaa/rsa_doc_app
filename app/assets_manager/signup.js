import { auth } from './firebase-config.js';
import { createUserWithEmailAndPassword, updateProfile, signOut } from 'https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js';

export async function createStaffAccount({ displayName, email, password }) {
  if (!auth) throw new Error('Account registration is not configured yet. Please contact your administrator.');
  const { user } = await createUserWithEmailAndPassword(auth, email, password);
  let profileSaved = true;
  try {
    await updateProfile(user, { displayName });
  } catch {
    // The account already exists: do not tell the user to submit it again.
    profileSaved = false;
  }
  await signOut(auth).catch(() => {});
  return { profileSaved };
}
