<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

// CORS for local dev + production domain.
$allowedOrigins = [
    'https://cmbankrsa.com',
    'http://127.0.0.1:5500',
    'http://localhost:5500',
    'http://127.0.0.1:8000',
    'http://localhost:8000'
];
$origin = $_SERVER['HTTP_ORIGIN'] ?? '';
if ($origin && in_array($origin, $allowedOrigins, true)) {
    header('Access-Control-Allow-Origin: ' . $origin);
    header('Vary: Origin');
}
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, Authorization');
if (($_SERVER['REQUEST_METHOD'] ?? '') === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if (!function_exists('str_contains')) {
    function str_contains(string $haystack, string $needle): bool {
        return $needle !== '' && mb_strpos($haystack, $needle) !== false;
    }
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Method not allowed']);
    exit;
}

// Load credentials from either:
// 1) api/backblaze-config.php (recommended on shared hosting), or
// 2) environment variables.
$applicationKeyId = '';
$applicationKey = '';
$bucketId = '';

$configPath = __DIR__ . DIRECTORY_SEPARATOR . 'backblaze-config.php';
if (file_exists($configPath)) {
    $cfg = require $configPath;
    if (is_array($cfg)) {
        $applicationKeyId = (string)($cfg['applicationKeyId'] ?? '');
        $applicationKey = (string)($cfg['applicationKey'] ?? '');
        $bucketId = (string)($cfg['bucketId'] ?? '');
    }
}

if (!$applicationKeyId) $applicationKeyId = (string)(getenv('BACKBLAZE_APPLICATION_KEY_ID') ?: '');
if (!$applicationKey) $applicationKey = (string)(getenv('BACKBLAZE_APPLICATION_KEY') ?: '');
if (!$bucketId) $bucketId = (string)(getenv('BACKBLAZE_BUCKET_ID') ?: '');

if (
    !$applicationKeyId ||
    !$applicationKey ||
    !$bucketId ||
    str_contains($applicationKeyId, 'PASTE') ||
    str_contains($applicationKey, 'PASTE') ||
    str_contains($bucketId, 'PASTE')
) {
    http_response_code(500);
    echo json_encode(['error' => 'Backblaze credentials are not configured on the server.']);
    exit;
}

function readJsonBody(): array {
    $raw = file_get_contents('php://input');
    if (!$raw) return [];
    $decoded = json_decode($raw, true);
    return is_array($decoded) ? $decoded : [];
}

function backblazeRequest(string $method, string $url, array $headers = [], ?string $body = null): array {
    $ch = curl_init($url);
    curl_setopt($ch, CURLOPT_CUSTOMREQUEST, $method);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    curl_setopt($ch, CURLOPT_TIMEOUT, 60);
    if ($body !== null) {
        curl_setopt($ch, CURLOPT_POSTFIELDS, $body);
    }

    $responseBody = curl_exec($ch);
    $curlErr = curl_error($ch);
    $status = (int) curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($responseBody === false) {
        return [0, ['error' => $curlErr ?: 'Network error']];
    }

    $decoded = json_decode($responseBody, true);
    if (!is_array($decoded)) {
        $decoded = ['raw' => $responseBody];
    }
    return [$status, $decoded];
}

try {
    $payload = isset($_FILES['file']) ? $_POST : readJsonBody();
    $action = (string) ($payload['action'] ?? '');

    if ($action === 'authorize') {
        $authHeader = 'Authorization: Basic ' . base64_encode($applicationKeyId . ':' . $applicationKey);
        [$status, $data] = backblazeRequest(
            'GET',
            'https://api.backblazeb2.com/b2api/v2/b2_authorize_account',
            [$authHeader, 'Accept: application/json']
        );

        if ($status < 200 || $status >= 300) {
            http_response_code($status ?: 502);
            echo json_encode(['error' => 'Backblaze authorization failed', 'details' => $data]);
            exit;
        }

        echo json_encode([
            'authorizationToken' => $data['authorizationToken'] ?? null,
            'apiUrl' => $data['apiUrl'] ?? null,
            'downloadUrl' => $data['downloadUrl'] ?? null,
            'bucketId' => $bucketId
        ]);
        exit;
    }

    if ($action === 'getUploadUrl') {
        $authorizationToken = (string) ($payload['authorizationToken'] ?? '');
        $apiUrl = (string) ($payload['apiUrl'] ?? '');
        $requestedBucketId = (string) ($payload['bucketId'] ?? '');

        if (!$authorizationToken || !$apiUrl || !$requestedBucketId) {
            http_response_code(400);
            echo json_encode(['error' => 'Missing authorizationToken, apiUrl, or bucketId']);
            exit;
        }

        if ($requestedBucketId !== $bucketId) {
            http_response_code(403);
            echo json_encode(['error' => 'Invalid bucket']);
            exit;
        }

        [$status, $data] = backblazeRequest(
            'POST',
            rtrim($apiUrl, '/') . '/b2api/v2/b2_get_upload_url',
            ['Authorization: ' . $authorizationToken, 'Content-Type: application/json'],
            json_encode(['bucketId' => $bucketId], JSON_UNESCAPED_SLASHES)
        );

        if ($status < 200 || $status >= 300) {
            http_response_code($status ?: 502);
            echo json_encode(['error' => 'Failed to get Backblaze upload URL', 'details' => $data]);
            exit;
        }

        echo json_encode([
            'uploadUrl' => $data['uploadUrl'] ?? null,
            'authorizationToken' => $data['authorizationToken'] ?? null
        ]);
        exit;
    }

    if ($action === 'upload') {
        if (!isset($_FILES['file']) || !is_uploaded_file($_FILES['file']['tmp_name'])) {
            http_response_code(400);
            echo json_encode(['error' => 'No file uploaded']);
            exit;
        }

        $uploadUrl = (string) ($payload['uploadUrl'] ?? '');
        $uploadAuthToken = (string) ($payload['uploadAuthToken'] ?? '');
        $fileName = (string) ($payload['fileName'] ?? '');
        $contentType = (string) ($payload['contentType'] ?? 'application/octet-stream');
        $sha1 = (string) ($payload['sha1'] ?? '');

        if (!$uploadUrl || !$uploadAuthToken || !$fileName || !$sha1) {
            http_response_code(400);
            echo json_encode(['error' => 'Missing upload parameters']);
            exit;
        }

        $fileBytes = file_get_contents($_FILES['file']['tmp_name']);
        if ($fileBytes === false) {
            http_response_code(500);
            echo json_encode(['error' => 'Failed to read uploaded file']);
            exit;
        }

        [$status, $data] = backblazeRequest(
            'POST',
            $uploadUrl,
            [
                'Authorization: ' . $uploadAuthToken,
                'X-Bz-File-Name: ' . rawurlencode($fileName),
                'Content-Type: ' . $contentType,
                'X-Bz-Content-Sha1: ' . $sha1,
                'X-Bz-Content-Disposition: inline'
            ],
            $fileBytes
        );

        if ($status < 200 || $status >= 300) {
            http_response_code($status ?: 502);
            echo json_encode(['error' => 'Backblaze upload failed', 'details' => $data]);
            exit;
        }

        echo json_encode([
            'fileId' => $data['fileId'] ?? null,
            'fileName' => $data['fileName'] ?? null
        ]);
        exit;
    }

    http_response_code(400);
    echo json_encode(['error' => 'Invalid action']);
} catch (Throwable $e) {
    http_response_code(500);
    echo json_encode(['error' => 'Server error', 'details' => $e->getMessage()]);
}
