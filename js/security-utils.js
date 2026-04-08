// ==================== SECURITY UTILITIES ====================
// This file provides security functions to protect against injection attacks

// Initialize security on page load
window.addEventListener('DOMContentLoaded', () => {
    initializeSecurityMeasures();
});

// Initialize all security measures
function initializeSecurityMeasures() {
    // Add Content Security Policy headers via meta tags
    addCSPHeaders();
    
    // Protect against direct DOM manipulation (non-blocking)
    protectDOMManipulation();

}

// ==================== CONTENT SECURITY POLICY ====================
function addCSPHeaders() {
    // Check if CSP meta tag already exists
    if (!document.querySelector('meta[http-equiv="Content-Security-Policy"]')) {
        const cspMeta = document.createElement('meta');
        cspMeta.httpEquiv = 'Content-Security-Policy';
        cspMeta.content = "default-src 'self'; script-src 'self' https://cdnjs.cloudflare.com https://cdn.jsdelivr.net https://www.gstatic.com https://unpkg.com https://www.googletagmanager.com 'unsafe-inline'; style-src 'self' https://cdnjs.cloudflare.com https://fonts.googleapis.com 'unsafe-inline'; font-src 'self' https://cdnjs.cloudflare.com https://fonts.gstatic.com data:; img-src 'self' data: https:; connect-src 'self' https://*.googleapis.com https://*.firebaseio.com https://*.firebaseio-demo.com https://*.gstatic.com https://*.firebaseapp.com https://*.cloudfunctions.net https://firebasestorage.googleapis.com https://securetoken.googleapis.com https://identitytoolkit.googleapis.com https://firebaseinstallations.googleapis.com https://firestore.googleapis.com https://storage.googleapis.com https://www.google-analytics.com https://www.googletagmanager.com wss://*.firebaseio.com wss://firestore.googleapis.com https://*.backblazeb2.com https://cors-proxy.naniadezz.workers.dev; frame-src 'self' data: blob: https:; child-src 'self' data: blob: https:;";
        document.head.appendChild(cspMeta);
    }
}

// ==================== SAFE HTML INSERTION ====================
// Use this function instead of innerHTML for user-generated content
function safeSetHTML(element, htmlContent) {
    if (!element) return;
    
    // Check if DOMPurify is available
    if (typeof DOMPurify !== 'undefined') {
        element.innerHTML = DOMPurify.sanitize(htmlContent, {
            ALLOWED_TAGS: ['b', 'i', 'em', 'strong', 'p', 'br', 'span', 'div'],
            ALLOWED_ATTR: ['class', 'id', 'style']
        });
    } else {
        // Fallback: use textContent instead
        element.textContent = htmlContent;
    }
}

// ==================== SAFE TEXT INSERTION ====================
function safeSetText(element, text) {
    if (!element) return;
    element.textContent = String(text || '');
}

// ==================== VALIDATE AND SANITIZE INPUT ====================
function sanitizeInput(input, type = 'text') {
    if (!input) return '';
    
    const str = String(input).trim();
    
    switch(type) {
        case 'email':
            // Basic email validation
            return str.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/) ? str : '';
        case 'phone':
            // Remove non-digit characters
            return str.replace(/\D/g, '');
        case 'number':
            // Only allow numbers and basic operators
            return str.replace(/[^0-9.,\-+*/()]/g, '');
        case 'text':
        default:
            // Remove potentially dangerous characters
            return str
                .replace(/<script[^>]*>.*?<\/script>/gi, '')
                .replace(/<iframe[^>]*>.*?<\/iframe>/gi, '')
                .replace(/javascript:/gi, '')
                .replace(/on\w+\s*=/gi, '')
                .substring(0, 1000); // Limit length
    }
}

// ==================== PROTECT DOM MANIPULATION ====================
function protectDOMManipulation() {
    // Note: We don't override innerHTML as it breaks Firebase and other libraries
    // Instead, use safeSetHTML() function when displaying user content
}

// ==================== MONITOR CONSOLE ACCESS ====================
const suspiciousPatterns = [
    /localStorage/gi,
    /sessionStorage/gi,
    /eval\(/gi,
    /Function\(/gi,
    /document\.cookie/gi
];

function monitorConsoleAccess() {
    // Monitor for suspicious patterns without blocking (for logging only)
    const originalLog = console.log;
    console.log = function(...args) {
        // In production, you can send logs to a server
        // For now, we just let it pass through
        originalLog.apply(console, args);
    };
    
    // DON'T override eval - it breaks Firebase auth
    // Just monitor for suspicious usage
}

// ==================== DISABLE CONSOLE (OPTIONAL) ====================
function disableConsoleInProduction() {
    // Only apply in production (not localhost)
    if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
        console.log = () => {};
        console.warn = () => {};
        console.error = () => {};
        console.info = () => {};
        console.debug = () => {};
    }
}

// ==================== VALIDATE FIREBASE OPERATIONS ====================
function validateFirebaseData(data) {
    if (!data) return null;
    
    // Ensure data is an object
    if (typeof data !== 'object') return null;
    
    // Remove any fields that might contain executable code
    const sanitized = {};
    for (const [key, value] of Object.entries(data)) {
        // Skip suspicious field names
        if (/script|eval|function|code|exec/i.test(key)) continue;
        
        // Sanitize string values
        if (typeof value === 'string') {
            sanitized[key] = sanitizeInput(value, 'text');
        } else if (typeof value === 'number' || typeof value === 'boolean') {
            sanitized[key] = value;
        } else if (Array.isArray(value)) {
            sanitized[key] = value.map(v => typeof v === 'string' ? sanitizeInput(v, 'text') : v);
        } else if (typeof value === 'object' && value !== null) {
            sanitized[key] = validateFirebaseData(value);
        }
    }
    
    return sanitized;
}

// ==================== EXPORT FUNCTIONS ====================
window.SecurityUtils = {
    safeSetHTML,
    safeSetText,
    sanitizeInput,
    validateFirebaseData,
    initializeSecurityMeasures
};
