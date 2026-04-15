# Admin Direct Password Reset (`123456`)

This project now supports admin direct reset to `123456` via backend API.

## What was added
- Admin button action in users table (`key` icon) now calls:
  - `POST /api/admin/reset-password`
- Endpoint implemented in:
  - `render-email-api/server.js`
- Frontend API base URL config:
  - `js/admin-api-config.js`

## Required configuration

1. Deploy/update backend service (`render-email-api`) on Render.
2. Set `FIREBASE_SERVICE_ACCOUNT_JSON` in Render env vars.
3. Ensure `REQUIRE_FIREBASE_AUTH=true` (recommended).
4. In frontend, set real backend URL in `js/admin-api-config.js`:

```js
export const ADMIN_API_BASE_URL = 'https://your-service.onrender.com';
```

## How it works
- Logged-in admin clicks reset password.
- Backend verifies Firebase ID token.
- Backend verifies caller role is admin from Firestore `users`.
- Backend sets target Firebase Auth password to `123456`.

## Notes
- This is a high-risk operation. Only admins should have access.
- Consider forcing user to change password after next login in future enhancement.
