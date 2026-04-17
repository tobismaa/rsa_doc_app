// js/backblaze-storage.js
const sharedBackblazeState = {
    session: null,
    uploadUrl: null
};

export class BackblazeStorage {
    constructor() {
        this.bucketName = 'cmbank-rsa-documents';
        const isLocalHost = ['127.0.0.1', 'localhost'].includes(window.location.hostname);
        this.apiProxyEndpoint = isLocalHost
            ? 'https://cmbankrsa.com/api/backblaze-upload.php'
            : '/api/backblaze-upload.php';
        this.bucketId = null;
        this.authorizationToken = null;
        this.apiUrl = null;
        this.downloadUrl = null;
        this.initFailed = false;
        this.initErrorMessage = '';
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
            throw error;
        }

        let data = null;
        try {
            data = await response.json();
        } catch (_) {
            data = null;
        }

        if (!response.ok) {
            const message = data?.error || `Request failed (${response.status})`;
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
            console.error('Initialization Error:', error);
            this.initFailed = true;
            this.initErrorMessage = String(error?.message || 'Initialization failed');
            throw error;
        }
    }

    async uploadFile(file, customerName, documentType) {
        try {
            if (!this.authorizationToken) {
                await this.init();
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
            console.error('Upload Error:', error);
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
