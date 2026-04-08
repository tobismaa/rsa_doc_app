// js/google-config.js
export const googleConfig = {
    clientId: 'REPLACE_WITH_GOOGLE_CLIENT_ID',
    clientSecret: 'REPLACE_WITH_GOOGLE_CLIENT_SECRET',
    redirectUris: [
        'http://localhost:8000',
        'http://localhost:63342'
    ],
    javascriptOrigins: [
        'http://localhost:8000',
        'http://localhost:63342'
    ],
    scopes: [
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive.metadata.readonly'
    ]
};
