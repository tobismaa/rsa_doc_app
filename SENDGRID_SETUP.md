# SendGrid Email Alert Setup (Firebase Trigger Email)

This project already queues email jobs into Firestore collection `mail` via `js/email-alerts.js`.
To actually deliver those emails, configure Firebase Extension `firestore-send-email` with SendGrid SMTP.

## 1. Create SendGrid sender + API key

1. In SendGrid, verify a sender identity (single sender or domain).
2. Create an API key with Mail Send permission.
3. Keep the full API key safe (you will use it as SMTP password).

## 2. Install / Configure Firebase extension

Install from Firebase Console or CLI:

```bash
firebase ext:install firebase/firestore-send-email --project=<your-project-id>
```

Use these values:

- `MAIL_COLLECTION`: `mail`
- `DEFAULT_FROM`: your verified sender email (for example `alerts@yourdomain.com`)
- `DEFAULT_REPLY_TO`: optional support email
- `SMTP_CONNECTION_URI`: `smtps://apikey@smtp.sendgrid.net:465`
- `SMTP_PASSWORD`: your SendGrid API key

Notes:
- SendGrid SMTP username is literally `apikey`.
- If your network blocks 465, use STARTTLS on 587 with: `smtp://apikey@smtp.sendgrid.net:587`.

## 3. Security rules (important)

Only trusted users/code should be able to write to `mail` documents, to prevent abuse.

At minimum, lock down direct client writes if possible, or only allow writes through controlled app paths/roles.

## 4. Verify end-to-end

1. Trigger a real app event (assignment, approval, rejection, etc.).
2. Confirm a new document appears in Firestore `mail`.
3. Confirm extension updates delivery metadata on that document.
4. Confirm recipient receives email.

## 5. Troubleshooting quick checks

- Auth errors (`535`): check username is exactly `apikey` and API key is complete.
- Connection errors: verify host is `smtp.sendgrid.net` and port is allowed.
- No email sent: confirm extension is installed in the same Firebase project (`rsa-doc-app`) and monitoring `mail` collection.

## Sources

- Firebase Trigger Email docs: https://firebase.google.com/docs/extensions/official/firestore-send-email
- Firebase Extension Hub (firestore-send-email): https://extensions.dev/extensions/firebase/firestore-send-email
- SendGrid SMTP auth/host guidance: https://support.sendgrid.com/hc/en-us/articles/17894803361819-Troubleshooting-535-Authentication-failed-Error
- SendGrid SMTP configuration examples: https://sendgrid.com/en-us/blog/how-to-migrate-from-mandrill-to-sendgrid
