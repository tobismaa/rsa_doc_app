// github-upload.js

const GITHUB_CONFIG = {
    username: 'tobismaa',
    repo: 'e-voucher-files',          // Updated to match your repo name
    branch: 'main',
    folder: 'vouchers'
};

// Helper: Convert File to Base64
const fileToBase64 = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = error => reject(error);
});

async function uploadToGitHub(file, onProgress) {
    const cleanName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_');
    const fileName = `${GITHUB_CONFIG.folder}/${Date.now()}_${cleanName}`;

    if (onProgress) onProgress(`Encoding ${file.name}...`);
    const content = await fileToBase64(file);

    const url = `https://api.github.com/repos/${GITHUB_CONFIG.username}/${GITHUB_CONFIG.repo}/contents/${fileName}`;

    const payload = {
        message: `Upload voucher: ${cleanName}`,
        content: content,
        branch: GITHUB_CONFIG.branch
    };

    if (onProgress) onProgress(`Uploading ${file.name}...`);

    try {
        const response = await fetch(url, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.message || 'Upload failed');
        }

        const data = await response.json();

        // Replace that line with this:
        const publicUrl = `https://cdn.jsdelivr.net/gh/${GITHUB_CONFIG.username}/${GITHUB_CONFIG.repo}@${GITHUB_CONFIG.branch}/${fileName}`;
        return {
            name: file.name,
            url: publicUrl,
            type: file.type
        };
    } catch (error) {
        console.error("GitHub Upload Error:", error);
        throw error;
    }
}

async function processGitHubUploads(filesArray, updateStatusFn) {
    const attachments = [];
    for (let i = 0; i < filesArray.length; i++) {
        const fileObj = filesArray[i];
        if (updateStatusFn) updateStatusFn(`Uploading... (${i + 1}/${filesArray.length})`);
        const result = await uploadToGitHub(fileObj.file, updateStatusFn);
        attachments.push(result);
    }
    return attachments;
}