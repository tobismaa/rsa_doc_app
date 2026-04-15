# EmailJS Setup (Simple Path)

This app now sends alerts using EmailJS directly from `js/email-alerts.js`.

## 1. Get your EmailJS values

From EmailJS dashboard:
- `Public Key`
- `Service ID`
- `Template ID`

## 2. Add values to config file

Open `js/emailjs-config.js` and paste your real values:

```js
export const EMAILJS_PUBLIC_KEY = 'your_public_key';
export const EMAILJS_SERVICE_ID = 'your_service_id';
export const EMAILJS_TEMPLATE_ID = 'your_template_id';
```

## 3. Create EmailJS template variables

In your EmailJS template, include these variables:
- `{{to_email}}`
- `{{email_subject}}`
- `{{email_message}}`
- `{{email_html}}`
- `{{notification_type}}`
- `{{submission_id}}`
- `{{customer_name}}`
- `{{event_key}}`

Minimum needed:
- `{{to_email}}`
- `{{email_subject}}`
- `{{email_message}}`

## 4. Test

Trigger one app workflow event (assignment/approval/rejection/final).

Then check Firestore `emailNotificationEvents`:
- `status: sent` means EmailJS accepted it
- `status: failed` means config/template issue
