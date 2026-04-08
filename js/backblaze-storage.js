// js/backblaze-storage.js
export class BackblazeStorage {
    constructor() {
        // ⚠️ CRITICAL: Replace these placeholders with your NEW keys generated today.
        // Do NOT use the old keys shared in the chat.
        this.applicationKeyId = '7cdc2e9d99b8';
        this.applicationKey = '005b13adad696e40439bba402cc2b93a696c4c28cd';
        
        this.bucketName = 'cmbank-rsa-documents';
        this.bucketId = '67fcddcc227ea9ed99c90b18';
        
        // 🚀 Using your own Cloudflare Worker (Durable & Secure)
        // Ensure your Worker allows 'http://127.0.0.1:5500'
        this.proxyBase = 'https://cors-proxy.naniadezz.workers.dev/?url=';
        
        // Base URLs (Cleaned: No trailing spaces)
        this.baseApiUrl = 'https://api.backblazeb2.com';
        this.baseDownloadUrl = 'https://f002.backblazeb2.com';
        
        this.authorizationToken = null;
        this.apiUrl = this.baseApiUrl;
        this.downloadUrl = this.baseDownloadUrl;
    }

    async init() {
        if (this.applicationKeyId.includes('PASTE')) {
            throw new Error('🛑 STOP: You must edit the file and insert your real Backblaze keys!');
        }

        try {
            console.log('🔄 Initializing Backblaze connection...');
            
            const authString = btoa(`${this.applicationKeyId}:${this.applicationKey}`);
            const authEndpoint = `${this.baseApiUrl}/b2api/v2/b2_authorize_account`;
            
            // Route through your proxy to bypass CORS
            const proxyUrl = `${this.proxyBase}${encodeURIComponent(authEndpoint)}`;

            const authResponse = await fetch(proxyUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Basic ${authString}`,
                    'Accept': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest'
                }
            });
            
            if (!authResponse.ok) {
                const errText = await authResponse.text();
                throw new Error(`Auth Failed (${authResponse.status}): ${errText}`);
            }
            
            const authData = await authResponse.json();
            
            // Update URLs based on response
            this.apiUrl = authData.apiUrl;
            this.downloadUrl = authData.downloadUrl;
            this.authorizationToken = authData.authorizationToken;
            
            console.log('✅ Backblaze Initialized Successfully');
            return true;
            
        } catch (error) {
            console.error('❌ Initialization Error:', error);
            throw error;
        }
    }

    async uploadFile(file, customerName, documentType) {
        try {
            if (!this.authorizationToken) {
                await this.init();
            }

            // 1. Get Upload URL
            const getUploadUrlEndpoint = `${this.apiUrl}/b2api/v2/b2_get_upload_url`;
            const proxyGetUrl = `${this.proxyBase}${encodeURIComponent(getUploadUrlEndpoint)}`;

            const urlResponse = await fetch(proxyGetUrl, {
                method: 'POST',
                headers: {
                    'Authorization': this.authorizationToken,
                    'Content-Type': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: JSON.stringify({ bucketId: this.bucketId })
            });

            if (!urlResponse.ok) throw new Error('Failed to get upload URL');
            const urlData = await urlResponse.json();
            
            // 2. Prepare File
            const safeName = (str) => str.replace(/[^a-zA-Z0-9._-]/g, '_');
            const fileName = `${Date.now()}_${safeName(customerName)}_${safeName(documentType)}_${safeName(file.name)}`;
            const sha1 = await this.calculateSHA1(file);

            // 3. Upload File
            // Note: The uploadUrl from Backblaze is direct. We proxy it to handle CORS.
            const proxyUploadUrl = `${this.proxyBase}${encodeURIComponent(urlData.uploadUrl)}`;

            const uploadResponse = await fetch(proxyUploadUrl, {
                method: 'POST',
                headers: {
                    'Authorization': urlData.authorizationToken,
                    'X-Bz-File-Name': encodeURIComponent(fileName),
                    'Content-Type': file.type || 'application/octet-stream',
                    'X-Bz-Content-Sha1': sha1,
                    'X-Bz-Content-Disposition': 'inline',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: file
            });

            if (!uploadResponse.ok) {
                const err = await uploadResponse.text();
                throw new Error(`Upload Failed: ${err}`);
            }

            const result = await uploadResponse.json();
            const publicUrl = `${this.downloadUrl}/file/${this.bucketName}/${encodeURIComponent(fileName)}`;

            console.log('✅ File Uploaded:', fileName);
            return {
                fileId: result.fileId,
                fileName: result.fileName,
                fileUrl: publicUrl,
                timestamp: Date.now()
            };

        } catch (error) {
            console.error('❌ Upload Error:', error);
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
                    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
                    resolve(hashHex);
                } catch (err) { reject(err); }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }
}