const crypto = require("crypto");
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");
const { S3Client, PutObjectCommand, GetObjectCommand } = require("@aws-sdk/client-s3");
const { getSignedUrl } = require("@aws-sdk/s3-request-presigner");

admin.initializeApp();

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new HttpsError("failed-precondition", `${name} is not configured on the server.`);
  }
  return value;
}

function getB2Client() {
  const endpoint = requireEnv("B2_ENDPOINT").replace(/\/+$/, "");
  const region = requireEnv("B2_REGION");
  const accessKeyId = requireEnv("B2_KEY_ID");
  const secretAccessKey = requireEnv("B2_APPLICATION_KEY");

  return new S3Client({
    region,
    endpoint,
    forcePathStyle: true,
    credentials: { accessKeyId, secretAccessKey }
  });
}

function cleanFileName(fileName) {
  const baseName = String(fileName || "attachment").split(/[\\/]/).pop();
  return baseName.replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 120) || "attachment";
}

function publicUrlFor(key) {
  const publicBase = process.env.B2_PUBLIC_BASE_URL;
  if (publicBase) {
    return `${publicBase.replace(/\/+$/, "")}/${encodeURI(key)}`;
  }

  const endpoint = requireEnv("B2_ENDPOINT").replace(/\/+$/, "");
  const bucket = requireEnv("B2_BUCKET");
  return `${endpoint}/${encodeURIComponent(bucket)}/${key.split("/").map(encodeURIComponent).join("/")}`;
}

exports.createB2UploadUrl = onCall(async (request) => {
  if (!request.auth) {
    throw new HttpsError("unauthenticated", "You must be logged in to upload voucher files.");
  }

  const bucket = requireEnv("B2_BUCKET");
  const fileName = cleanFileName(request.data && request.data.fileName);
  const contentType = String((request.data && request.data.contentType) || "application/octet-stream");
  const size = Number(request.data && request.data.size);

  if (!Number.isFinite(size) || size <= 0) {
    throw new HttpsError("invalid-argument", "A valid file size is required.");
  }

  const randomId = crypto.randomBytes(8).toString("hex");
  const key = `vouchers/${request.auth.uid}/${Date.now()}-${randomId}-${fileName}`;
  const client = getB2Client();
  const command = new PutObjectCommand({
    Bucket: bucket,
    Key: key,
    ContentType: contentType
  });

  const uploadUrl = await getSignedUrl(client, command, { expiresIn: 10 * 60 });

  return {
    uploadUrl,
    key,
    url: publicUrlFor(key),
    provider: "backblaze_b2",
    expiresInSeconds: 10 * 60
  };
});

exports.createB2DownloadUrl = onCall(async (request) => {
  if (!request.auth) {
    throw new HttpsError("unauthenticated", "You must be logged in to view voucher files.");
  }

  const bucket = requireEnv("B2_BUCKET");
  const key = String(request.data && request.data.key || "");
  if (!key || !key.startsWith("vouchers/")) {
    throw new HttpsError("invalid-argument", "A valid voucher file key is required.");
  }

  const client = getB2Client();
  const command = new GetObjectCommand({ Bucket: bucket, Key: key });
  const downloadUrl = await getSignedUrl(client, command, { expiresIn: 10 * 60 });

  return { downloadUrl, expiresInSeconds: 10 * 60 };
});
