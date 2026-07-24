// js/backblaze-storage.js
import { db } from './firebase-config.js';
import { getSystemSettings } from './shared/system-settings.js?v=20260724a';

const STORAGE_CAP_FULL_UPLOAD_MESSAGE = 'Failed to upload. Storage cap full.';

const sharedBackblazeState = {
    session: null,
    uploadUrl: null
};

export class BackblazeStorage {
    constructor() {
        this.bucketName = 'cmbank-rsa-documents';
        const defaultRenderEndpoint = 'https://rsa-email-api.onrender.com/api/backblaze-upload';
        this.apiProxyEndpoint = String(window.__BACKBLAZE_UPLOAD_API_URL__ || defaultRenderEndpoint).trim();
        this.bucketId = null;
        this.authorizationToken = null;
        this.apiUrl = null;
        this.downloadUrl = null;
        this.initFailed = false;
        this.initErrorMessage = '';
    }

    formatProxyError(response, data) {
        const detail = data?.details || {};
        const detailMessage = detail?.message || detail?.error || detail?.code || '';
        const baseMessage = data?.error || `Request failed (${response.status})`;
        return detailMessage && detailMessage !== baseMessage
            ? `${baseMessage}: ${detailMessage}`
            : baseMessage;
    }

    async callProxy(action, payload = null, file = null) {
        let response;

        try {
            if (file) {
                const formData = new FormData();
                formData.append('action', action);
                Object.entries(payload || {}).forEach(([key, value]) => {
                    if (value !== undefined && value !== null) {
                        formData.append(key, String(value));
                    }
                });
                formData.append('file', file);

                response = await fetch(this.apiProxyEndpoint, {
                    method: 'POST',
                    body: formData
                });
            } else {
                response = await fetch(this.apiProxyEndpoint, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action, ...(payload || {}) })
                });
            }
        } catch (error) {
            if (['127.0.0.1', 'localhost'].includes(window.location.hostname)) {
                throw new Error(`Cannot reach upload API at ${this.apiProxyEndpoint}.`);
            }
            throw new Error(`Cannot reach upload API. Check your internet connection and try again. (${error?.message || 'Network error'})`);
        }

        let data = null;
        try {
            data = await response.json();
        } catch (_) {
            data = null;
        }

        if (!response.ok) {
            const message = this.formatProxyError(response, data);
            if (response.status === 405) {
                throw new Error('Upload API endpoint rejected POST (405).');
            }
            throw new Error(message);
        }

        return data || {};
    }

    async init() {
        if (this.initFailed) {
            throw new Error(this.initErrorMessage || 'Backblaze initialization previously failed.');
        }
        if (this.authorizationToken) return true;

        if (sharedBackblazeState.session) {
            this.authorizationToken = sharedBackblazeState.session.authorizationToken;
            this.apiUrl = sharedBackblazeState.session.apiUrl;
            this.downloadUrl = sharedBackblazeState.session.downloadUrl;
            this.bucketId = sharedBackblazeState.session.bucketId;
            return true;
        }

        try {
            const authData = await this.callProxy('authorize');

            this.apiUrl = authData.apiUrl;
            this.downloadUrl = authData.downloadUrl;
            this.authorizationToken = authData.authorizationToken;
            this.bucketId = authData.bucketId;

            if (!this.authorizationToken || !this.bucketId) {
                throw new Error('Server did not return Backblaze authorization data.');
            }

            sharedBackblazeState.session = {
                authorizationToken: this.authorizationToken,
                apiUrl: this.apiUrl,
                downloadUrl: this.downloadUrl,
                bucketId: this.bucketId
            };

            return true;
        } catch (error) {
            this.initFailed = true;
            this.initErrorMessage = String(error?.message || 'Initialization failed');
            throw error;
        }
    }

    async assertUploadsEnabled() {
        const settings = await getSystemSettings(db, { force: true });
        if (settings?.uploadControls?.simulateStorageCapFullFailure === true) {
            throw new Error(STORAGE_CAP_FULL_UPLOAD_MESSAGE);
        }
    }

    async uploadFile(file, customerName, documentType) {
        try {
            await this.assertUploadsEnabled();
            if (!this.authorizationToken) {
                await this.init();
            }
            if (!file) {
                throw new Error('No file selected for upload.');
            }

            let uploadUrlData = sharedBackblazeState.uploadUrl;
            if (!uploadUrlData) {
                uploadUrlData = await this.callProxy('getUploadUrl', {
                    authorizationToken: this.authorizationToken,
                    apiUrl: this.apiUrl,
                    bucketId: this.bucketId
                });
                sharedBackblazeState.uploadUrl = uploadUrlData;
            }
            if (!uploadUrlData?.uploadUrl || !uploadUrlData?.authorizationToken) {
                sharedBackblazeState.uploadUrl = null;
                throw new Error('Upload server did not return a valid Backblaze upload URL.');
            }

            const safeName = (str) => String(str || '').replace(/[^a-zA-Z0-9._-]/g, '_');
            const fileName = `${Date.now()}_${safeName(customerName)}_${safeName(documentType)}_${safeName(file.name)}`;
            const sha1 = await this.calculateSHA1(file);

            let uploadResult;
            try {
                uploadResult = await this.callProxy(
                    'upload',
                    {
                        uploadUrl: uploadUrlData.uploadUrl,
                        uploadAuthToken: uploadUrlData.authorizationToken,
                        fileName,
                        contentType: file.type || 'application/octet-stream',
                        sha1
                    },
                    file
                );
            } catch (error) {
                sharedBackblazeState.uploadUrl = null;
                uploadUrlData = await this.callProxy('getUploadUrl', {
                    authorizationToken: this.authorizationToken,
                    apiUrl: this.apiUrl,
                    bucketId: this.bucketId
                });
                sharedBackblazeState.uploadUrl = uploadUrlData;
                uploadResult = await this.callProxy(
                    'upload',
                    {
                        uploadUrl: uploadUrlData.uploadUrl,
                        uploadAuthToken: uploadUrlData.authorizationToken,
                        fileName,
                        contentType: file.type || 'application/octet-stream',
                        sha1
                    },
                    file
                );
            }

            const publicUrl = `${this.downloadUrl}/file/${this.bucketName}/${encodeURIComponent(fileName)}`;

            return {
                fileId: uploadResult.fileId,
                fileName: uploadResult.fileName || fileName,
                fileUrl: publicUrl,
                timestamp: Date.now()
            };
        } catch (error) {
            console.error('Backblaze upload failed', error);
            throw error;
        }
    }

    async calculateSHA1(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const buffer = e.target.result;
                    const hashBuffer = await crypto.subtle.digest('SHA-1', buffer);
                    const hashArray = Array.from(new Uint8Array(hashBuffer));
                    const hashHex = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
                    resolve(hashHex);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }
}
