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
    CURLOPT_FOLLOWLOCATION => true,
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

header('Content-Type: ' . ($contentType ?: 'application/octet-stream'));
header('Content-Disposition: attachment; filename="' . addslashes($fileName) . '"');
header('Content-Length: ' . strlen($body));
header('Cache-Control: private, max-age=300');
echo $body;
