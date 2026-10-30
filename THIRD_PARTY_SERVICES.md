# RSA Application: Third-Party Services

## Scope

This inventory lists the third-party services used by the RSA application and explains the role each service plays in the app.

## Active or directly wired services

| Service | What it does in the application | Status / evidence | Payment range coverage | Amount |
|---|---|---|---|---|
| **Firebase Authentication** | Handles user sign-in, sign-out, session validation, role access, and Firebase-based password resets. | Active through `js/firebase-config.js` and `js/auth.js`. | Covered under the Firebase Blaze Plan through April 2027 | **$480/year ($40/month)** — combined Blaze Plan covering Authentication, Cloud Firestore, FCM, and Analytics |
| **Cloud Firestore** | Stores and synchronizes users, applications, document metadata, roles, settings, reports, chat data, and workflow state. | Active through the Firebase SDK and dashboard modules. | Covered under the Firebase Blaze Plan through April 2027 | |
| **Firebase Cloud Messaging (FCM)** | Registers browser push tokens and displays foreground/background notifications for chat and application events. | Active when the browser supports notifications and the user grants permission. | Covered under the Firebase Blaze Plan through April 2027 | |
| **Firebase Analytics** | Collects application analytics when supported. | Enabled only outside local development/file mode and when browser support is available. | Covered under the Firebase Blaze Plan through April 2027 | |
| **Render** (`rsa-email-api.onrender.com`) | Hosts the backend API used for account verification, administrative actions, document-upload proxying, push registration, and email-related endpoints. | Directly configured in `js/email-api-config.js`, `js/admin-api-config.js`, and `js/backblaze-storage.js`. | One-year coverage; renewable October 2027 | **$300/year — Yet to be paid** |
| **Backblaze B2** | Stores up to **500 GB** of uploaded RSA documents and provides document download URLs. | Active through the Render upload proxy and `js/backblaze-storage.js`. | 500 GB storage; payment coverage through January 2029 | Amount not provided |
| **Paystack** | Provides Nigerian bank lists and verifies bank account names during agent/customer registration workflows. | Active through the Render backend; the browser calls `/api/paystack/banks` and `/api/paystack/resolve-account`. | Account verification; free for now | **Free for now** |
| **EmailJS** | Sends application notification emails and scheduled report emails. | Direct REST calls exist in `js/email-alerts.js` and `render-email-api/server.js`. Requires valid EmailJS configuration. | 2,000 email messages per month; payment coverage through April 2027 | **$108/year + $8.01 VAT** |
| **WhatsApp** | Opens WhatsApp conversations from phone-number links in the dashboards. | Used only when a user clicks a generated WhatsApp link; there is no WhatsApp API integration. | Contact messaging; free for now | **Free for now** |
| **Hostinger Domain** | Provides the RSA application domain registration. | Added as the application domain provider. | Domain coverage from April 14, 2026 to April 14, 2028 | **₦61,800 for 2 years** |
| **Hostinger Hosting** | Hosts the RSA application website and related files. | Added as the application hosting provider. | Hosting coverage through April 14, 2030 | **₦475,200 for 4 years** |

## External libraries and delivery networks

These providers deliver frontend libraries or fonts rather than business data services:

| Provider | Libraries/assets used |
|---|---|
| **Google Fonts** | Inter font |
| **Cloudflare cdnjs** | Font Awesome, jsPDF, PDF.js, html2canvas, jsPDF AutoTable |
| **jsDelivr** | DOMPurify, ExcelJS, PizZip, SheetJS |
| **unpkg** | pdf-lib and fontkit |
| **Google `gstatic.com`** | Firebase browser SDK modules |

## Main service flow

```text
RSA dashboards
    |
    +--> Firebase Authentication / Firestore / FCM / Analytics
    |
    +--> Render backend API
              |
              +--> Backblaze B2 document storage
              +--> Paystack bank/account verification
              +--> EmailJS email delivery
```

