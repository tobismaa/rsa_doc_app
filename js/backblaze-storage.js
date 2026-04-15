// js/backblaze-storage.js
export class BackblazeStorage {
    constructor() {
        this.bucketName = 'cmbank-rsa-documents';
        const runtimeOverride = String(window.__BACKBLAZE_API_ENDPOINT__ || '').trim();
        this.apiProxyEndpoint = runtimeOverride || this.resolveDefaultEndpoint();
        this.bucketId = null;
        this.authorizationToken = null;
        this.apiUrl = null;
        this.downloadUrl = null;
        this.initFailed = false;
        this.initErrorMessage = '';
    }

    resolveDefaultEndpoint() {
        const { protocol, hostname, port, origin } = window.location;
        const isLocal = hostname === '127.0.0.1' || hostname === 'localhost';
        const isStaticDevServer = isLocal && (port === '5500' || protocol === 'file:');

        if (isStaticDevServer) {
            const localOrigin = protocol === 'file:'
                ? 'http://127.0.0.1:8000'
                : origin.replace(/:\d+$/, ':8000');
            return `${localOrigin}/api/backblaze-upload.php`;
        }

        return '/api/backblaze-upload.php';
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
            const isLocal = window.location.hostname === '127.0.0.1' || window.location.hostname === 'localhost';
            if (isLocal) {
                throw new Error(`Cannot reach upload API at ${this.apiProxyEndpoint}. Start a PHP server for this project, for example: php -S 127.0.0.1:8000`);
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
                throw new Error(`Upload API rejected POST at ${this.apiProxyEndpoint}. Use a PHP runtime or set window.__BACKBLAZE_API_ENDPOINT__ to a live server endpoint.`);
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

        try {
            console.log('Initializing Backblaze connection via server endpoint...');
            const authData = await this.callProxy('authorize');

            this.apiUrl = authData.apiUrl;
            this.downloadUrl = authData.downloadUrl;
            this.authorizationToken = authData.authorizationToken;
            this.bucketId = authData.bucketId;

            if (!this.authorizationToken || !this.bucketId) {
                throw new Error('Server did not return Backblaze authorization data.');
            }

            console.log('Backblaze initialized successfully');
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

            const uploadUrlData = await this.callProxy('getUploadUrl', {
                authorizationToken: this.authorizationToken,
                apiUrl: this.apiUrl,
                bucketId: this.bucketId
            });

            const safeName = (str) => String(str || '').replace(/[^a-zA-Z0-9._-]/g, '_');
            const fileName = `${Date.now()}_${safeName(customerName)}_${safeName(documentType)}_${safeName(file.name)}`;
            const sha1 = await this.calculateSHA1(file);

            const uploadResult = await this.callProxy(
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

            const publicUrl = `${this.downloadUrl}/file/${this.bucketName}/${encodeURIComponent(fileName)}`;

            console.log('File uploaded:', fileName);
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
