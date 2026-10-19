const { test } = require('node:test');
const assert = require('node:assert/strict');
const { createDocumentPreviewHandler } = require('./document-preview');
const url = 'https://f005.backblazeb2.com/file/cmbank-rsa-documents/sample.pdf';
function result() {
    return { code: 200, headers: {}, set(k,v) { this.headers[k] = v; return this; },
        status(code) { this.code = code; return this; }, type(value) { return this.set('Content-Type',value); },
        send(body) { this.body = body; return this; } };
}
test('PDF returned inline, with no-store and correct MIME even if upstream uses octet-stream', async () => {
    const res = result();
    await createDocumentPreviewHandler(async (target, options) => {
        assert.equal(target, url);
        assert.equal(options.redirect, 'error');
        return new Response('%PDF-1.7 test', { headers: { 'Content-Type': 'application/octet-stream' } });
    })({ query: { url } },res);
    assert.equal(res.code,200);
    assert.equal(res.headers['Content-Type'],'application/pdf');
    assert.match(res.headers['Content-Disposition'],/^inline;/);
    assert.equal(res.headers['Cache-Control'],'private, no-store');
    assert.equal(res.body.toString(),'%PDF-1.7 test');
});
test('invalid storage URLs rejected before network access', async () => {
    for (const value of ['https://evil.example/a', url.replace('cmbank-rsa-documents','other'), url.replace('https:', 'http:'), url.replace('f005.', 'f005.backblazeb2.com.evil.'), url.replace('f005.', 'user:pass@f005.'), url.replace('/sample.pdf','/../../other/file.pdf')]) {
        const res = result();
        await createDocumentPreviewHandler(() => { throw new Error('Must not fetch'); })({query: {url:value}},res);
        // A normalized path within this bucket is still allowed; traversal out is rejected.
        assert.equal(res.code,400);
    }
});
test('HTML cannot be rendered as a document', async () => {
    const res = result();
    await createDocumentPreviewHandler(async () => new Response('<html>unsafe</html>'))({query:{url}},res);
    assert.equal(res.code,415);
});
test('storage errors and failed fetches produce visible preview errors', async () => {
    for (const fetcher of [async () => new Response('missing',{status:404}), async () => { throw new Error('timeout'); }]) {
        const res = result();
        await createDocumentPreviewHandler(fetcher)({query:{url}},res);
        assert.equal(res.code,502);
        assert.match(res.body,/document|Document/);
    }
});

test('download mode returns attachments and preserves bytes for non-preview formats', async () => {
    const bytes = Buffer.from('PK example office document');
    const res = result();
    await createDocumentPreviewHandler(async () => new Response(bytes))({query:{url,download:'1',filename:'report.docx'}},res);
    assert.equal(res.code,200);
    assert.equal(res.headers['Content-Type'],'application/octet-stream');
    assert.equal(res.headers['Content-Disposition'],'attachment; filename="report.docx"');
    assert.deepEqual(res.body,bytes);
});
test('download mode sanitizes names and preserves PDF content type', async () => {
    const res = result();
    await createDocumentPreviewHandler(async () => new Response('%PDF-1.7 test'))({query:{url,download:'1',filename:'bad"/name.pdf'}},res);
    assert.equal(res.headers['Content-Disposition'],'attachment; filename="bad__name.pdf"');
    assert.equal(res.headers['Content-Type'],'application/pdf');
});
