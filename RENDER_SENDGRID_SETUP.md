# Render + SendGrid Email Alert Setup

This setup removes dependency on Firebase Trigger Email extension.
Your frontend now calls a Render API, and Render sends email through SendGrid.

## What changed in this repo

- Frontend email module now posts to `/api/send-alert`:
  - `js/email-alerts.js`
- Render API base URL config:
  - `js/email-api-config.js`
- Render backend service:
  - `render-email-api/server.js`
  - `render-email-api/package.json`
  - `render-email-api/.env.example`

## 1. Create Render Web Service

1. Push this repo to GitHub.
2. In Render:
   - `New` -> `Web Service`
   - Connect your repo.
   - Set:
     - `Root Directory`: `render-email-api`
     - `Build Command`: `npm install`
     - `Start Command`: `npm start`
     - `Environment`: `Node`

## 2. Add environment variables in Render

From `render-email-api/.env.example`, set:

- `SENDGRID_API_KEY`: your SendGrid API key (Mail Send permission)
- `SENDGRID_FROM_EMAIL`: verified sender (for example `alerts@yourdomain.com`)
- `REQUIRE_FIREBASE_AUTH`: `true`
- `ALLOWED_ORIGINS`: comma-separated origins for your app
  - example: `http://localhost:8000,http://localhost:63342,https://your-frontend-domain.com`
- `FIREBASE_SERVICE_ACCOUNT_JSON`: your full Firebase service account JSON on one line

Notes:
- Keep `REQUIRE_FIREBASE_AUTH=true` in production.
- In Firebase Console -> Project Settings -> Service Accounts, generate key JSON.

## 3. Point frontend to Render

Edit `js/email-api-config.js`:

```js
export const EMAIL_API_BASE_URL = 'https://your-render-service.onrender.com';
```

## 4. Deploy frontend update

Deploy your frontend files with the updated API URL.

## 5. Verify

1. Open:
   - `https://your-render-service.onrender.com/health`
2. Expected:
   - `{ "ok": true, ... }`
3. In app, trigger one workflow event (assignment/approval/rejection/final).
4. Confirm recipient inbox receives alert.
5. Confirm Firestore `emailNotificationEvents/<eventKey>` status updates to `sent`.

## 6. Troubleshooting

- `401 Missing Firebase ID token`
  - user is not authenticated, or auth header is missing.
- `CORS` error
  - frontend origin not listed in `ALLOWED_ORIGINS`.
- `Unsupported email type`
  - payload `meta.type` must be one of supported workflow types.
- `SENDGRID_FROM_EMAIL` rejected
  - sender identity not verified in SendGrid.
