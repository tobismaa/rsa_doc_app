// Serve previews through the Node API; static dashboard hosting cannot execute PHP.
export function getDocumentPreviewUrl(value) {
    const cleanUrl = String(value || '').trim();
    if (!cleanUrl) return '';
    try {
        const url = new URL(cleanUrl, window.location.origin);
        if (url.protocol === 'https:' && /\.backblazeb2\.com$/i.test(url.hostname)
            && url.pathname.startsWith('/file/cmbank-rsa-documents/')) {
            const uploadEndpoint = new URL(window.__BACKBLAZE_UPLOAD_API_URL__ || 'https://rsa-email-api.onrender.com/api/backblaze-upload');
            const previewEndpoint = new URL('/api/document-preview', uploadEndpoint.origin);
            previewEndpoint.searchParams.set('url', url.href);
            return previewEndpoint.href;
        }
        return cleanUrl;
    } catch (_) {
        return '';
    }
}
