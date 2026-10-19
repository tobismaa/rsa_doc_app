const MAX_BYTES = 50 * 1024 * 1024;

function documentType(bytes) {
    if (bytes.subarray(0, 5).toString() === '%PDF-') return 'application/pdf';
    if (bytes.subarray(0, 8).equals(Buffer.from([137,80,78,71,13,10,26,10]))) return 'image/png';
    if (bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255) return 'image/jpeg';
    if (['GIF87a', 'GIF89a'].includes(bytes.subarray(0, 6).toString())) return 'image/gif';
    if (bytes.subarray(0, 4).toString() === 'RIFF' && bytes.subarray(8, 12).toString() === 'WEBP') return 'image/webp';
    return '';
}

function createDocumentPreviewHandler(fetchDocument = fetch) {
    return async (req, res) => {
        res.set('Cache-Control', 'private, no-store');
        res.set('X-Content-Type-Options', 'nosniff');
        let url;
        try {
            url = new URL(String(req.query.url || ''));
            if (url.protocol !== 'https:' || !/^f[0-9]+\.backblazeb2\.com$/i.test(url.hostname)
                || url.port || url.username || url.password
                || !url.pathname.startsWith('/file/cmbank-rsa-documents/')) throw new Error();
        } catch (_) {
            return res.status(400).type('text/plain').send('Invalid document URL.');
        }
        try {
            const response = await fetchDocument(url.href, {
                redirect: 'error', signal: AbortSignal.timeout(120000)
            });
            if (!response.ok) {
                await response.body?.cancel();
                return res.status(502).type('text/plain').send('Unable to retrieve this document from storage. Please try again.');
            }
            const reader = response.body.getReader();
            const chunks = [];
            let total = 0;
            try {
                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;
                    total += value.byteLength;
                    if (total > MAX_BYTES) {
                        await reader.cancel();
                        return res.status(413).type('text/plain').send('Document exceeds the 50 MB preview limit.');
                    }
                    chunks.push(Buffer.from(value));
                }
            } finally {
                reader.releaseLock();
            }
            const bytes = Buffer.concat(chunks);
            const download = req.query.download === '1';
            const type = documentType(bytes);
            if (!type && !download) return res.status(415).type('text/plain').send('This file type cannot be previewed. Download the document to view it.');
            const filename = String((download && req.query.filename) || decodeURIComponent(url.pathname.split('/').pop())).replace(/[^A-Za-z0-9._ -]/g, '_') || 'document';
            res.set('Content-Type', type || 'application/octet-stream');
            res.set('Content-Disposition', `${download ? 'attachment' : 'inline'}; filename="${filename}"`);
            res.set('Content-Security-Policy', "script-src 'none'");
            return res.send(bytes);
        } catch (_) {
            return res.status(502).type('text/plain').send('Document preview is temporarily unavailable. Please try again.');
        }
    };
}

module.exports = { createDocumentPreviewHandler };
