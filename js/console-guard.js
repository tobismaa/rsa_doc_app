// Silence console output in production to reduce exposed runtime details.
(function () {
    const host = String(window.location.hostname || '').toLowerCase();
    const isLocalhost = host === 'localhost' || host === '127.0.0.1';
    if (isLocalhost) return;

    const noop = function () {};
    try {
        console.log = noop;
        console.warn = noop;
        console.error = noop;
        console.info = noop;
        console.debug = noop;
        console.trace = noop;
    } catch (_) {
        // ignore
    }
})();
