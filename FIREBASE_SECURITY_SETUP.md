# Firebase Security Setup

This project now includes a stricter Firestore ruleset:

- `firestore.rules`

## Important pre-step before deploying rules

The rules read role info from `users/{uid}`.

If your existing users were created with random document IDs, migrate them so each signed-in user has a profile document at:

- `users/<firebase-auth-uid>`

## Minimum migration plan

1. For each existing user profile document in `users`:
- Keep the same data.
- Create/merge a new doc with id = that profile's `uid`.

2. Ensure at least one admin account has:
- `role: "admin"` (or `"super_admin"`)
- `status: "active"`

3. After migration, publish rules from Firebase console or CLI:
- Firebase Console -> Firestore Database -> Rules -> paste `firestore.rules` contents -> Publish.

## New signup behavior

`js/auth.js` now writes new users directly to `users/{uid}`, which aligns with the rules model.

## Notes

- If rules are published before migration, some existing users may get `permission-denied`.
- `audit` writes are append-only (no updates/deletes from client).
- `settings` writes are restricted to `super_admin`.
