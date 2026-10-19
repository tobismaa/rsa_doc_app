<?php
declare(strict_types=1);

if ($_SERVER['REQUEST_METHOD'] !== 'GET') {
    http_response_code(405);
    header('Allow: GET');
    echo 'Method not allowed';
    exit;
}

$url = trim((string)($_GET['url'] ?? ''));
if ($url === '') {
    http_response_code(400);
    echo 'Missing document URL';
    exit;
}

$parts = parse_url($url);
$host = strtolower((string)($parts['host'] ?? ''));
$path = (string)($parts['path'] ?? '');
$scheme = strtolower((string)($parts['scheme'] ?? ''));

$isBackblazeHost = $host === 'backblazeb2.com' || substr($host, -strlen('.backblazeb2.com')) === '.backblazeb2.com';
$isExpectedBucket = strpos($path, '/file/cmbank-rsa-documents/') === 0;

if ($scheme !== 'https' || !$isBackblazeHost || !$isExpectedBucket) {
    http_response_code(400);
    echo 'Invalid document URL';
    exit;
}

$fileName = basename(rawurldecode($path)) ?: 'document.pdf';
$fileName = preg_replace('/[^A-Za-z0-9._ -]+/', '_', $fileName) ?: 'document.pdf';

$ch = curl_init($url);
curl_setopt_array($ch, [
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_FOLLOWLOCATION => false,
    CURLOPT_CONNECTTIMEOUT => 15,
    CURLOPT_TIMEOUT => 120,
    CURLOPT_HEADER => false,
    CURLOPT_SSL_VERIFYPEER => true,
    CURLOPT_USERAGENT => 'CMBankRSA-DocumentProxy/1.0',
]);

$body = curl_exec($ch);
$status = (int)curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
$contentType = (string)curl_getinfo($ch, CURLINFO_CONTENT_TYPE);
$error = curl_error($ch);
curl_close($ch);

if ($body === false || $status < 200 || $status >= 300) {
    http_response_code(502);
    echo 'Document download failed' . ($error ? ': ' . $error : '');
    exit;
}

$inline = ($_GET['inline'] ?? '') === '1';
if ($inline) {
    // Inspect bytes rather than trusting the remote Content-Type for same-origin content.
    $detector = new finfo(FILEINFO_MIME_TYPE);
    $contentType = $detector->buffer($body);
    $allowedTypes = ['application/pdf', 'image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    if (!in_array($contentType, $allowedTypes, true)) {
        http_response_code(415);
        header('Content-Type: text/plain; charset=utf-8');
        echo 'This file type cannot be previewed. Download the document to view it.';
        exit;
    }
}
header('X-Content-Type-Options: nosniff');
header("Content-Security-Policy: script-src 'none'; frame-ancestors 'self'");
header('Content-Type: ' . ($contentType ?: 'application/octet-stream'));
header('Content-Disposition: ' . ($inline ? 'inline' : 'attachment') . '; filename="' . addslashes($fileName) . '"');
header('Content-Length: ' . strlen($body));
header('Cache-Control: private, no-store');
echo $body;
