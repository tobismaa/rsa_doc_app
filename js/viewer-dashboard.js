// js/viewer-dashboard.js - COMPLETE WORKING VERSION WITH FIXED SAVE
import { auth, db } from './firebase-config.js';
import {
    collection,
    query,
    where,
    orderBy,
    onSnapshot,
    updateDoc,
    doc,
    getDocs
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

// ==================== CORS PROXY CONFIG ====================
const CORS_PROXY = 'https://cors-proxy.naniadezz.workers.dev?url=';

// ==================== DOCUMENT TYPES MAPPING ====================
const DOCUMENT_TYPES = {
    'birth_certificate': 'Birth Certificate',
    'nin': 'National ID (NIN)',
    'pay_slips': 'Pay Slips',
    'offer_letter': 'Offer Letter',
    'intro_letter': 'Introduction Letter',
    'request_letter': 'Request Letter',
    'rsa_statement': 'RSA Statement',
    'pfa_form': 'PFA Form',
    'consent_letter': 'Consent Letter',
    'indemnity_form': 'Indemnity Form',
    'utility_bill': 'Utility Bill',
    'benefit_application_form': 'Benefit Application Form'
};

// ==================== GLOBAL VARIABLES ====================
let currentUser = null;
let currentUserData = null;
let allSubmissions = [];
let currentSubmissionId = null;
let currentViewerSubmission = null;
let currentViewerIndex = 0;
let currentMergeSubmission = null;
let mergeSelectionOrder = [];
let mergedPdfBlob = null;
let mergedPdfUrl = null;
let downloadInProgress = false;
let mergeInProgress = false;
let additionalFiles = [];
let additionalFileCounter = 0;

// Cache for user full names
const userFullNameCache = new Map();

// ==================== DOM ELEMENTS ====================
const userName = document.getElementById('userName');
const userAvatar = document.getElementById('userAvatar');
const pendingTableBody = document.getElementById('pendingTableBody');
const approvedTableBody = document.getElementById('approvedTableBody');
const rejectedTableBody = document.getElementById('rejectedTableBody');
const commentModal = document.getElementById('commentModal');
const mergeModal = document.getElementById('mergeModal');
const viewerModal = document.getElementById('viewerModal');
const closeCommentModal = document.getElementById('closeCommentModal');
const closeMergeModal = document.getElementById('closeMergeModal');
const closeViewer = document.getElementById('closeViewer');
const cancelComment = document.getElementById('cancelComment');
const cancelMerge = document.getElementById('cancelMerge');
const approveBtn = document.getElementById('approveDocument');
const rejectBtn = document.getElementById('rejectDocument');
const mergeDocumentsBtn = document.getElementById('mergeDocumentsBtn');
const commentText = document.getElementById('commentText');
const modalCustomerName = document.getElementById('modalCustomerName');
const modalDocumentType = document.getElementById('modalDocumentType');
const viewerFileName = document.getElementById('viewerFileName');
const documentViewer = document.getElementById('documentViewer');
const viewerSaveBtn = document.getElementById('viewerSaveBtn');
const viewerDownloadSection = document.getElementById('viewerDownloadSection');
const viewerMergedSaveBtn = document.getElementById('viewerMergedSaveBtn');
const viewerNav = document.getElementById('viewerNav');
const mergeList = document.getElementById('mergeList');
const mergeCustomerName = document.getElementById('mergeCustomerName');
const mergeDownloadSection = document.getElementById('mergeDownloadSection');
const notification = document.getElementById('notification');
const pageTitle = document.getElementById('pageTitle');
const pendingCountBadge = document.getElementById('pendingCount');
const zoomInBtn = document.getElementById('zoomInBtn');
const zoomOutBtn = document.getElementById('zoomOutBtn');
const zoomResetBtn = document.getElementById('zoomResetBtn');
const zoomLevel = document.getElementById('zoomLevel');
const selectedCount = document.getElementById('selectedCount');

// ==================== ZOOM LEVEL ====================
let currentZoom = 1.0;

// ==================== CORS HELPER FUNCTION ====================
async function fetchWithCorsFallback(url) {
    const cleanUrl = url?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) throw new Error('Invalid URL');

zoom    // Skip direct fetch attempt for Backblaze URLs to avoid CORS console errors
    if (cleanUrl.includes('backblazeb2.com')) {
        const proxyUrl = `${CORS_PROXY}${encodeURIComponent(cleanUrl)}`;
        const response = await fetch(proxyUrl);
        if (!response.ok) throw new Error(`Proxy fetch failed: ${response.status}`);
        return response;
    }

    // Try direct fetch first
    try {
        const response = await fetch(cleanUrl, {
            mode: 'cors',
            credentials: 'omit',
            headers: {
                'Accept': 'application/pdf, image/*, */*'
            }
        });
        if (response.ok) return response;
    } catch (e) {
        console.log('Direct fetch failed, trying proxy...');
    }

    // Use proxy
    const proxyUrl = `${CORS_PROXY}${encodeURIComponent(cleanUrl)}`;
    const response = await fetch(proxyUrl);

    if (!response.ok) {
        throw new Error(`Proxy fetch failed: ${response.status}`);
    }
    return response;
}

// ==================== GET USER FULL NAME BY EMAIL ====================
async function getUserFullName(email) {
    if (!email) return 'Unknown';

    if (userFullNameCache.has(email)) {
        return userFullNameCache.get(email);
    }

    try {
        const userQuery = query(collection(db, 'users'), where('email', '==', email));
        const userSnapshot = await getDocs(userQuery);

        if (!userSnapshot.empty) {
            const userData = userSnapshot.docs[0].data();
            const fullName = userData.fullName || userData.displayName || email.split('@')[0];
            userFullNameCache.set(email, fullName);
            return fullName;
        }
    } catch (err) {
        console.warn('Could not fetch user name for:', email);
    }

    const fallbackName = email.split('@')[0];
    userFullNameCache.set(email, fallbackName);
    return fallbackName;
}

// ==================== CHECK IF FILE IS PDF ====================
function isPdfFile(fileUrl, fileName) {
    const lowerName = (fileName || '').toLowerCase();
    const lowerUrl = (fileUrl || '').toLowerCase();
    return lowerName.endsWith('.pdf') || lowerUrl.includes('.pdf') || lowerUrl.includes('application/pdf');
}

// ==================== CHECK IF FILE IS IMAGE ====================
function isImageFile(fileName) {
    const lowerName = fileName.toLowerCase();
    return lowerName.endsWith('.jpg') ||
           lowerName.endsWith('.jpeg') ||
           lowerName.endsWith('.png') ||
           lowerName.endsWith('.gif') ||
           lowerName.endsWith('.bmp') ||
           lowerName.endsWith('.webp');
}

// ==================== CONVERT IMAGE TO PDF ====================
async function imageToPdf(imageBlob, imageName) {
    try {
        const { PDFDocument } = PDFLib;
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([595, 842]); // A4 size

        const imageBytes = await imageBlob.arrayBuffer();
        let image;
        const lowerName = imageName.toLowerCase();

        // Try to detect image type
        try {
            if (lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg') || imageBlob.type.includes('jpeg')) {
                image = await pdfDoc.embedJpg(imageBytes);
            } else if (lowerName.endsWith('.png') || imageBlob.type.includes('png')) {
                image = await pdfDoc.embedPng(imageBytes);
            } else {
                // Check magic numbers
                const arr = new Uint8Array(imageBytes.slice(0, 4));
                if (arr[0] === 0xFF && arr[1] === 0xD8) {
                    image = await pdfDoc.embedJpg(imageBytes);
                } else if (arr[0] === 0x89 && arr[1] === 0x50 && arr[2] === 0x4E && arr[3] === 0x47) {
                    image = await pdfDoc.embedPng(imageBytes);
                } else {
                    throw new Error(`Unsupported image format: ${imageName}`);
                }
            }
        } catch (embedError) {
            console.error('Embed error:', embedError);
            throw new Error(`Failed to embed image: ${embedError.message}`);
        }

        const pageWidth = page.getWidth();
        const pageHeight = page.getHeight();
        const imgWidth = image.width;
        const imgHeight = image.height;
        const scale = Math.min(pageWidth / imgWidth, pageHeight / imgHeight) * 0.9;

        page.drawImage(image, {
            x: (pageWidth - imgWidth * scale) / 2,
            y: (pageHeight - imgHeight * scale) / 2,
            width: imgWidth * scale,
            height: imgHeight * scale,
        });

        return await pdfDoc.save();
    } catch (error) {
        console.error('Image to PDF conversion error:', error);
        throw new Error(`Failed to convert image to PDF: ${error.message}`);
    }
}

// ==================== SAVE FILE WITH LOCATION PICKER ====================
async function saveFileWithLocationPicker(blob, defaultFileName) {
    // Check if showSaveFilePicker is supported
    if ('showSaveFilePicker' in window) {
        try {
            const fileHandle = await window.showSaveFilePicker({
                suggestedName: defaultFileName,
                types: [{
                    description: 'PDF Document',
                    accept: { 'application/pdf': ['.pdf'] }
                }]
            });

            const writable = await fileHandle.createWritable();
            await writable.write(blob);
            await writable.close();

            showNotification(`✅ Saved: ${defaultFileName}`, 'success');
            return true;
        } catch (error) {
            if (error.name === 'AbortError') {
                showNotification('Save cancelled', 'info');
                return false;
            } else {
                console.error('Save error:', error);
                // Fall back to direct download
                return downloadBlobAsFile(blob, defaultFileName);
            }
        }
    } else {
        // Fallback for browsers without showSaveFilePicker
        return downloadBlobAsFile(blob, defaultFileName);
    }
}

// ==================== DOWNLOAD BLOB AS FILE (FALLBACK) ====================
function downloadBlobAsFile(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showNotification('✅ Download started', 'success');
    return true;
}

// ==================== FORMAT TIMESTAMP ====================
function formatTimestamp(timestamp) {
    if (!timestamp) return 'N/A';

    try {
        const date = timestamp.toDate ? timestamp.toDate() : new Date(timestamp);
        return date.toLocaleString('en-NG', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    } catch (e) {
        return 'N/A';
    }
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', () => {
    auth.onAuthStateChanged(async (user) => {
        if (user) {
            currentUser = user;

            try {
                const userQuery = query(collection(db, 'users'), where('email', '==', user.email));
                const userSnapshot = await getDocs(userQuery);

                if (!userSnapshot.empty) {
                    currentUserData = userSnapshot.docs[0].data();
                    userName.textContent = currentUserData.fullName || currentUserData.displayName || user.email.split('@')[0];
                } else {
                    userName.textContent = user.displayName || user.email.split('@')[0];
                }
            } catch (err) {
                console.error('Error fetching user data:', err);
                userName.textContent = user.displayName || user.email.split('@')[0];
            }

            userAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3E👤%3C/text%3E%3C/svg%3E';

            const q = query(collection(db, 'users'), where('email', '==', user.email));
            const snapshot = await getDocs(q);

            if (!snapshot.empty && snapshot.docs[0].data().role === 'viewer') {
                loadSubmissions();
            } else {
                showNotification('Access denied. Viewer privileges required.', 'error');
                setTimeout(() => { window.location.href = 'index.html'; }, 2000);
            }
        } else {
            window.location.href = 'index.html';
        }
    });
    setupEventListeners();
    setupZoomControls();
});

// ==================== SETUP ZOOM CONTROLS ====================
function setupZoomControls() {
    if (zoomInBtn) {
        zoomInBtn.addEventListener('click', () => {
            currentZoom = Math.min(currentZoom + 0.25, 3.0);
            applyZoom();
        });
    }

    if (zoomOutBtn) {
        zoomOutBtn.addEventListener('click', () => {
            currentZoom = Math.max(currentZoom - 0.25, 0.5);
            applyZoom();
        });
    }

    if (zoomResetBtn) {
        zoomResetBtn.addEventListener('click', () => {
            currentZoom = 1.0;
            applyZoom();
        });
    }
}

function applyZoom() {
    if (documentViewer) {
        documentViewer.style.transform = `scale(${currentZoom})`;
        documentViewer.style.transformOrigin = 'center top';
        documentViewer.style.transition = 'transform 0.3s ease';
    }
    if (zoomLevel) {
        zoomLevel.textContent = `${Math.round(currentZoom * 100)}%`;
    }
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
    if (closeCommentModal) closeCommentModal.addEventListener('click', () => closeModal(commentModal));
    if (closeMergeModal) closeMergeModal.addEventListener('click', () => closeMergeModalFunc());
    if (closeViewer) closeViewer.addEventListener('click', closeViewerModal);
    if (cancelComment) cancelComment.addEventListener('click', () => closeModal(commentModal));
    if (cancelMerge) cancelMerge.addEventListener('click', () => closeMergeModalFunc());

    if (approveBtn) approveBtn.addEventListener('click', () => reviewDocument('approved'));
    if (rejectBtn) rejectBtn.addEventListener('click', () => reviewDocument('rejected'));
    if (mergeDocumentsBtn) mergeDocumentsBtn.addEventListener('click', handleMergeDocuments);

    if (viewerSaveBtn) viewerSaveBtn.addEventListener('click', saveCurrentDocument);
    if (viewerMergedSaveBtn) viewerMergedSaveBtn.addEventListener('click', saveMergedPDF);

    window.addEventListener('click', (e) => {
        if (e.target === commentModal) closeModal(commentModal);
        if (e.target === mergeModal) closeMergeModalFunc();
        if (e.target === viewerModal) closeViewerModal();
    });
}

// ==================== LOAD SUBMISSIONS ====================
async function loadSubmissions() {
    const q = query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc'));

    onSnapshot(q, async (snapshot) => {
        allSubmissions = [];

        const uploaderEmails = [...new Set(snapshot.docs.map(doc => doc.data().uploadedBy))];
        const reviewerEmails = [...new Set(snapshot.docs.map(doc => doc.data().reviewedBy))];
        const allEmails = [...new Set([...uploaderEmails, ...reviewerEmails].filter(Boolean))];

        await Promise.all(allEmails.map(email => getUserFullName(email)));

        for (const docSnap of snapshot.docs) {
            const data = docSnap.data();
            const uploaderName = await getUserFullName(data.uploadedBy);

            let reviewerName = null;
            if (data.reviewedBy) {
                reviewerName = await getUserFullName(data.reviewedBy);
            }

            allSubmissions.push({
                id: docSnap.id,
                ...data,
                uploadedByName: uploaderName,
                reviewedByName: reviewerName
            });
        }

        renderAllTables();
        updatePendingCount();
    });
}

// ==================== UPDATE PENDING COUNT ====================
function updatePendingCount() {
    const pendingSubmissions = allSubmissions.filter(sub => sub.status === 'pending').length;
    if (pendingCountBadge) {
        pendingCountBadge.textContent = pendingSubmissions;
    }
}

// ==================== RENDER TABLES ====================
function renderAllTables() {
    const pendingSubs = allSubmissions.filter(sub => sub.status === 'pending');
    renderPendingTable(pendingSubs);

    const approvedSubs = allSubmissions.filter(sub => sub.status === 'approved');
    renderApprovedTable(approvedSubs);

    const rejectedSubs = allSubmissions.filter(sub => sub.status === 'rejected');
    renderRejectedTable(rejectedSubs);
}

function renderPendingTable(submissions) {
    if (!pendingTableBody) return;

    if (submissions.length === 0) {
        pendingTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No pending documents found</td></tr>';
        return;
    }

    pendingTableBody.innerHTML = submissions.map(sub => {
        let date = 'N/A';
        if (sub.uploadedAt) {
            date = formatTimestamp(sub.uploadedAt);
        }

        const docTypes = sub.documentTypes?.map(type => DOCUMENT_TYPES[type] || type).join(', ') || 'N/A';
        const docCount = sub.documents?.length || 0;

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${date}</td>
                <td>${docTypes} <small class="text-muted">(${docCount})</small></td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td><span class="status-badge status-pending">Pending</span></td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')">
                            <i class="fas fa-info-circle"></i> Details
                        </button>
                        <button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')">
                            <i class="fas fa-eye"></i> View
                        </button>
                        <button class="action-btn download-all-btn" onclick="window.downloadAll('${sub.id}')" ${downloadInProgress ? 'disabled' : ''}>
                            ${downloadInProgress ? '<i class="fas fa-spinner fa-spin"></i>' : '<i class="fas fa-download"></i>'} Download All
                        </button>
                        <button class="action-btn review-btn" onclick="window.openReviewModal('${sub.id}')">
                            <i class="fas fa-check-circle"></i> Review
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function renderApprovedTable(submissions) {
    if (!approvedTableBody) return;

    if (submissions.length === 0) {
        approvedTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No approved documents found</td></tr>';
        return;
    }

    approvedTableBody.innerHTML = submissions.map(sub => {
        let uploadDate = 'N/A';
        if (sub.uploadedAt) {
            uploadDate = formatTimestamp(sub.uploadedAt);
        }

        let approvedDate = 'N/A';
        if (sub.reviewedAt) {
            approvedDate = formatTimestamp(sub.reviewedAt);
        }

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td>${sub.reviewedByName || 'N/A'}</td>
                <td>${approvedDate}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')">
                            <i class="fas fa-info-circle"></i> Details
                        </button>
                        <button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')">
                            <i class="fas fa-eye"></i> View
                        </button>
                        <button class="action-btn download-all-btn" onclick="window.downloadAll('${sub.id}')" ${downloadInProgress ? 'disabled' : ''}>
                            ${downloadInProgress ? '<i class="fas fa-spinner fa-spin"></i>' : '<i class="fas fa-download"></i>'} Download All
                        </button>
                        <button class="action-btn merge-btn-small" onclick="window.openMergeModal('${sub.id}')" ${mergeInProgress ? 'disabled' : ''}>
                            <i class="fas fa-compress-alt"></i> Merge
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function renderRejectedTable(submissions) {
    if (!rejectedTableBody) return;

    if (submissions.length === 0) {
        rejectedTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No rejected documents found</td></tr>';
        return;
    }

    rejectedTableBody.innerHTML = submissions.map(sub => {
        let date = 'N/A';
        if (sub.uploadedAt) {
            date = formatTimestamp(sub.uploadedAt);
        }

        const docTypes = sub.documentTypes?.map(type => DOCUMENT_TYPES[type] || type).join(', ') || 'N/A';
        const docCount = sub.documents?.length || 0;

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${date}</td>
                <td>${docTypes} <small class="text-muted">(${docCount})</small></td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td>${sub.comment || 'No reason provided'}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')">
                            <i class="fas fa-info-circle"></i> Details
                        </button>
                        <button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')">
                            <i class="fas fa-eye"></i> View
                        </button>
                        <button class="action-btn download-all-btn" onclick="window.downloadAll('${sub.id}')" ${downloadInProgress ? 'disabled' : ''}>
                            ${downloadInProgress ? '<i class="fas fa-spinner fa-spin"></i>' : '<i class="fas fa-download"></i>'} Download All
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

// ==================== VIEW SUBMISSION ====================
window.viewSubmission = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub || !sub.documents || sub.documents.length === 0) {
        showNotification('No documents available', 'error');
        return;
    }

    // Store current submission and reset index
    currentViewerSubmission = sub;
    currentViewerIndex = 0;

    if (viewerSaveBtn) {
        viewerSaveBtn.style.display = 'inline-flex';
    }

    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'none';
    }

    // Show the first document
    showDocumentAtIndex(0);
};

function showDocumentAtIndex(index) {
    if (!currentViewerSubmission || !currentViewerSubmission.documents ||
        index < 0 || index >= currentViewerSubmission.documents.length) {
        return;
    }

    const doc = currentViewerSubmission.documents[index];
    const docTypeLabel = DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';

    // Update filename with counter if multiple documents
    if (currentViewerSubmission.documents.length > 1) {
        viewerFileName.textContent = `${docTypeLabel} (${index + 1}/${currentViewerSubmission.documents.length})`;
    } else {
        viewerFileName.textContent = docTypeLabel;
    }

    // Set document source
    documentViewer.src = doc.fileUrl?.trim();

    // Store current document data for saving
    if (viewerSaveBtn) {
        viewerSaveBtn.dataset.currentUrl = doc.fileUrl?.trim();
        viewerSaveBtn.dataset.currentName = doc.name || 'document.pdf';
        viewerSaveBtn.dataset.customerName = currentViewerSubmission.customerName || 'Customer';
        viewerSaveBtn.dataset.docType = docTypeLabel;
    }

    // Update navigation buttons
    updateViewerNavigation(index);

    // Show modal and reset zoom
    viewerModal.classList.add('active');
    currentZoom = 1.0;
    applyZoom();
}

function updateViewerNavigation(currentIndex) {
    if (!viewerNav || !currentViewerSubmission) return;

    const totalDocs = currentViewerSubmission.documents.length;

    if (totalDocs <= 1) {
        viewerNav.innerHTML = '';
        return;
    }

    viewerNav.innerHTML = `
        <button class="nav-btn" id="prevDocBtn" ${currentIndex === 0 ? 'disabled' : ''}>
            <i class="fas fa-chevron-left"></i>
        </button>
        <span class="nav-counter">${currentIndex + 1}/${totalDocs}</span>
        <button class="nav-btn" id="nextDocBtn" ${currentIndex === totalDocs - 1 ? 'disabled' : ''}>
            <i class="fas fa-chevron-right"></i>
        </button>
    `;

    const prevBtn = document.getElementById('prevDocBtn');
    const nextBtn = document.getElementById('nextDocBtn');

    if (prevBtn) {
        prevBtn.addEventListener('click', () => {
            if (currentViewerIndex > 0) {
                currentViewerIndex--;
                showDocumentAtIndex(currentViewerIndex);
            }
        });
    }

    if (nextBtn) {
        nextBtn.addEventListener('click', () => {
            if (currentViewerIndex < currentViewerSubmission.documents.length - 1) {
                currentViewerIndex++;
                showDocumentAtIndex(currentViewerIndex);
            }
        });
    }
}

// ==================== DOWNLOAD ALL ====================
window.downloadAll = async (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    if (downloadInProgress) {
        showNotification('Download already in progress', 'warning');
        return;
    }

    try {
        downloadInProgress = true;

        if (!sub.documents || sub.documents.length === 0) {
            showNotification('No documents to download', 'error');
            return;
        }

        // Download each document
        for (const [index, doc] of sub.documents.entries()) {
            const docTypeLabel = DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';

            showNotification(`📄 Downloading ${index + 1}/${sub.documents.length}: ${docTypeLabel}...`, 'info');

            try {
                const response = await fetchWithCorsFallback(doc.fileUrl?.trim());
                const blob = await response.blob();

                // Generate filename
                const safeCustomerName = sub.customerName.replace(/[^a-zA-Z0-9\s_-]/g, '_').trim();
                const fileName = `${safeCustomerName}_${docTypeLabel}.pdf`;

                // Use save file picker for each file
                await saveFileWithLocationPicker(blob, fileName);
            } catch (err) {
                console.error(`Failed to download ${docTypeLabel}:`, err);
                showNotification(`⚠️ Failed: ${docTypeLabel}`, 'warning');
            }
        }

    } catch (error) {
        console.error('Download error:', error);
        showNotification('Download failed: ' + error.message, 'error');
    } finally {
        downloadInProgress = false;
    }
};

// ==================== OPEN REVIEW MODAL ====================
window.openReviewModal = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    currentSubmissionId = submissionId;
    modalCustomerName.textContent = sub.customerName;
    modalDocumentType.textContent = `${sub.documents?.length || 0} documents uploaded`;
    commentText.value = '';
    commentModal.classList.add('active');
};

// ==================== REVIEW DOCUMENT ====================
async function reviewDocument(action) {
    if (!currentSubmissionId) return;

    const comment = commentText.value.trim();

    if (action === 'rejected' && !comment) {
        showNotification('Please provide a reason for rejection', 'error');
        return;
    }

    try {
        const submissionRef = doc(db, 'submissions', currentSubmissionId);
        await updateDoc(submissionRef, {
            status: action,
            comment: comment || '',
            reviewedBy: currentUser.email,
            reviewedAt: new Date().toISOString()
        });

        showNotification(`Document ${action === 'approved' ? 'approved' : 'rejected'} successfully!`, 'success');
        closeModal(commentModal);

    } catch (error) {
        console.error('Review error:', error);
        showNotification('Failed to update status', 'error');
    }
}

// ==================== VIEW DOCUMENT ====================
window.viewDocument = (fileUrl, fileName, originalName = '') => {
    viewerFileName.textContent = fileName;
    documentViewer.src = fileUrl?.trim();

    if (viewerSaveBtn) {
        viewerSaveBtn.style.display = 'inline-flex';
        viewerSaveBtn.dataset.currentUrl = fileUrl?.trim();
        viewerSaveBtn.dataset.currentName = originalName || 'document.pdf';
    }

    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'none';
    }

    // Clear navigation for single document view
    if (viewerNav) {
        viewerNav.innerHTML = '';
    }

    viewerModal.classList.add('active');
    currentZoom = 1.0;
    applyZoom();
};

function closeViewerModal() {
    viewerModal.classList.remove('active');
    documentViewer.src = '';
    currentZoom = 1.0;
    currentViewerSubmission = null;
    currentViewerIndex = 0;

    if (viewerSaveBtn) {
        viewerSaveBtn.style.display = 'none';
    }

    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'none';
    }

    if (viewerNav) {
        viewerNav.innerHTML = '';
    }
}

// ==================== SAVE CURRENT DOCUMENT ====================
async function saveCurrentDocument() {
    if (!viewerSaveBtn || !viewerSaveBtn.dataset.currentUrl) {
        showNotification('No document to save', 'error');
        return;
    }

    const fileUrl = viewerSaveBtn.dataset.currentUrl;
    const customerName = viewerSaveBtn.dataset.customerName || 'Customer';
    const docType = viewerSaveBtn.dataset.docType || 'Document';

    try {
        showNotification('📄 Fetching document...', 'info');
        const response = await fetchWithCorsFallback(fileUrl);
        const blob = await response.blob();

        // Generate a clean filename
        const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s_-]/g, '_').trim();
        const fileName = `${safeCustomerName}_${docType}.pdf`;

        await saveFileWithLocationPicker(blob, fileName);
    } catch (error) {
        console.error('Save error:', error);
        showNotification('Save failed: ' + error.message, 'error');
    }
}

// ==================== ADDITIONAL FILES UPLOAD ====================
function setupAdditionalFileUpload() {
    const uploadArea = document.getElementById('additionalUploadArea');
    const fileInput = document.getElementById('additionalFileInput');

    if (!uploadArea || !fileInput) return;

    // Remove existing listeners by cloning
    const newUploadArea = uploadArea.cloneNode(true);
    uploadArea.parentNode.replaceChild(newUploadArea, uploadArea);

    const newFileInput = fileInput.cloneNode(true);
    fileInput.parentNode.replaceChild(newFileInput, fileInput);

    const updatedUploadArea = document.getElementById('additionalUploadArea');
    const updatedFileInput = document.getElementById('additionalFileInput');

    updatedUploadArea.addEventListener('click', () => {
        updatedFileInput.click();
    });

    updatedUploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        updatedUploadArea.style.borderColor = 'var(--cm-primary)';
        updatedUploadArea.style.background = 'var(--cm-light)';
    });

    updatedUploadArea.addEventListener('dragleave', () => {
        updatedUploadArea.style.borderColor = '#cbd5e1';
        updatedUploadArea.style.background = 'transparent';
    });

    updatedUploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        updatedUploadArea.style.borderColor = '#cbd5e1';
        updatedUploadArea.style.background = 'transparent';

        const files = Array.from(e.dataTransfer.files);
        handleAdditionalFiles(files);
    });

    updatedFileInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        handleAdditionalFiles(files);
    });
}

function handleAdditionalFiles(files) {
    if (files.length + additionalFiles.length > 10) {
        showNotification('Maximum 10 files total allowed', 'error');
        return;
    }

    // Support common image formats
    const allowedTypes = ['.pdf', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'];
    const validFiles = files.filter(file => {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        return allowedTypes.includes(ext);
    });

    if (validFiles.length !== files.length) {
        showNotification('Some files were skipped (only PDF and common image formats allowed)', 'warning');
    }

    validFiles.forEach(file => {
        const fileId = `additional_${additionalFileCounter++}`;
        additionalFiles.push({
            id: fileId,
            file: file,
            name: file.name,
            type: file.type,
            size: file.size,
            isAdditional: true
        });
    });

    refreshMergeList();
    updateFilePreview();
    updateSelectedCount();
}

function updateFilePreview() {
    const previewContainer = document.getElementById('additionalFilesPreview');
    if (!previewContainer) return;

    if (additionalFiles.length === 0) {
        previewContainer.innerHTML = '';
        return;
    }

    let html = `
        <div class="preview-header">
            <span><i class="fas fa-paperclip"></i> Added Files (${additionalFiles.length})</span>
            <button class="clear-files-btn" onclick="window.clearAdditionalFiles()">
                <i class="fas fa-trash"></i> Clear All
            </button>
        </div>
    `;

    additionalFiles.forEach(file => {
        const fileExt = '.' + file.name.split('.').pop().toLowerCase();
        const isPdf = fileExt === '.pdf';
        const fileSize = (file.size / 1024).toFixed(1) + ' KB';

        html += `
            <div class="preview-item">
                <i class="fas ${isPdf ? 'fa-file-pdf' : 'fa-file-image'}" style="color: ${isPdf ? '#dc2626' : '#10b981'};"></i>
                <span class="file-name" title="${file.name}">${file.name.substring(0, 25)}${file.name.length > 25 ? '...' : ''}</span>
                <span class="file-size">${fileSize}</span>
                ${!isPdf ? '<span class="convert-badge">→PDF</span>' : ''}
                <button class="remove-file-btn" onclick="window.removeAdditionalFile('${file.id}')" title="Remove">
                    <i class="fas fa-times-circle"></i>
                </button>
            </div>
        `;
    });

    previewContainer.innerHTML = html;
}

window.removeAdditionalFile = (fileId) => {
    additionalFiles = additionalFiles.filter(f => f.id !== fileId);
    mergeSelectionOrder = mergeSelectionOrder.filter(id => id !== fileId);
    refreshMergeList();
    updateFilePreview();
    renumberBadges();
    updateSelectedCount();
};

window.clearAdditionalFiles = () => {
    additionalFiles = [];
    mergeSelectionOrder = mergeSelectionOrder.filter(id => !id.startsWith('additional_'));
    refreshMergeList();
    updateFilePreview();
    const fileInput = document.getElementById('additionalFileInput');
    if (fileInput) fileInput.value = '';
    renumberBadges();
    updateSelectedCount();
};

// ==================== MERGE FUNCTIONS ====================
window.openMergeModal = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    currentMergeSubmission = sub;
    mergeSelectionOrder = [];
    additionalFiles = [];
    additionalFileCounter = 0;
    mergedPdfBlob = null;
    if (mergedPdfUrl) {
        URL.revokeObjectURL(mergedPdfUrl);
        mergedPdfUrl = null;
    }

    if (mergeDownloadSection) {
        mergeDownloadSection.style.display = 'none';
    }

    mergeCustomerName.textContent = sub.customerName;

    if (!sub.documents || sub.documents.length === 0) {
        showNotification('No documents available to merge', 'error');
        return;
    }

    refreshMergeList();

    const previewContainer = document.getElementById('additionalFilesPreview');
    if (previewContainer) {
        previewContainer.innerHTML = '';
    }

    const fileInput = document.getElementById('additionalFileInput');
    if (fileInput) fileInput.value = '';

    setTimeout(setupAdditionalFileUpload, 100);

    mergeModal.classList.add('active');
    updateSelectedCount();
};

window.handleMergeCheckbox = (checkbox, itemId) => {
    if (checkbox.checked) {
        if (!mergeSelectionOrder.includes(itemId)) {
            mergeSelectionOrder.push(itemId);
        }
    } else {
        mergeSelectionOrder = mergeSelectionOrder.filter(id => id !== itemId);
    }

    renumberBadges();
    updateSelectedCount();

    // Update select all checkbox
    const selectAllCheckbox = document.getElementById('selectAllDocuments');
    if (selectAllCheckbox) {
        const totalItems = document.querySelectorAll('.merge-item input[type="checkbox"]').length;
        const checkedItems = mergeSelectionOrder.length;
        selectAllCheckbox.checked = checkedItems === totalItems && totalItems > 0;
        selectAllCheckbox.indeterminate = checkedItems > 0 && checkedItems < totalItems;
    }
};

window.selectAllDocuments = (checkbox) => {
    const checkboxes = document.querySelectorAll('.merge-item input[type="checkbox"]');
    mergeSelectionOrder = [];

    if (checkbox.checked) {
        checkboxes.forEach(cb => {
            cb.checked = true;
            const itemId = cb.value;
            mergeSelectionOrder.push(itemId);
        });
    } else {
        checkboxes.forEach(cb => {
            cb.checked = false;
        });
    }

    renumberBadges();
    updateSelectedCount();
};

function updateSelectedCount() {
    if (selectedCount) {
        selectedCount.textContent = mergeSelectionOrder.length;
    }
}

function renumberBadges() {
    mergeSelectionOrder.forEach((itemId, order) => {
        const badge = document.getElementById(`order_${itemId}`);
        if (badge) {
            badge.textContent = `#${order + 1}`;
            badge.style.display = 'inline-block';
        }
    });

    // Hide badges for unchecked items
    document.querySelectorAll('.merge-order-badge').forEach(badge => {
        if (!mergeSelectionOrder.includes(badge.id.replace('order_', ''))) {
            badge.textContent = '';
            badge.style.display = 'none';
        }
    });
}

function refreshMergeList() {
    if (!currentMergeSubmission) return;

    const originalDocs = currentMergeSubmission.documents || [];

    let combinedHtml = `
        <div class="merge-list-header">
            <div class="select-all">
                <input type="checkbox" id="selectAllDocuments" onchange="window.selectAllDocuments(this)">
                <label for="selectAllDocuments">Select All</label>
            </div>
            <div class="merge-instructions">
                <i class="fas fa-info-circle"></i>
                <span>Check documents to merge. Order will be preserved (#1, #2, #3...)</span>
            </div>
        </div>
        <div class="merge-items-container">
    `;

    // Original documents
    originalDocs.forEach((doc, index) => {
        const docTypeLabel = DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';
        const isPdf = isPdfFile(doc.fileUrl, doc.name);
        const itemId = `orig_${index}`;
        const isChecked = mergeSelectionOrder.includes(itemId);

        combinedHtml += `
            <div class="merge-item ${isChecked ? 'selected' : ''}" data-type="original" data-index="${index}" data-url="${doc.fileUrl}" data-name="${doc.name}">
                <input type="checkbox" id="merge_${itemId}" value="${itemId}" onchange="window.handleMergeCheckbox(this, '${itemId}')" ${isChecked ? 'checked' : ''}>
                <label for="merge_${itemId}" class="merge-item-label">
                    <i class="fas ${isPdf ? 'fa-file-pdf' : 'fa-file-image'}" style="color: ${isPdf ? '#dc2626' : '#10b981'};"></i>
                    <span class="doc-type">${docTypeLabel}</span>
                    <span class="doc-filename">(${doc.name || 'File'})</span>
                    ${!isPdf ? '<span class="merge-convert-badge">📷 Convert to PDF</span>' : ''}
                </label>
                <span class="merge-order-badge" id="order_${itemId}" style="display: ${isChecked ? 'inline-block' : 'none'};">${isChecked ? '#' + (mergeSelectionOrder.indexOf(itemId) + 1) : ''}</span>
            </div>
        `;
    });

    // Additional files
    additionalFiles.forEach((file) => {
        const fileExt = '.' + file.name.split('.').pop().toLowerCase();
        const isPdf = fileExt === '.pdf';
        const itemId = file.id;
        const isChecked = mergeSelectionOrder.includes(itemId);

        combinedHtml += `
            <div class="merge-item additional-item ${isChecked ? 'selected' : ''}" data-type="additional" data-file-id="${itemId}" data-name="${file.name}">
                <input type="checkbox" id="merge_${itemId}" value="${itemId}" onchange="window.handleMergeCheckbox(this, '${itemId}')" ${isChecked ? 'checked' : ''}>
                <label for="merge_${itemId}" class="merge-item-label">
                    <i class="fas ${isPdf ? 'fa-file-pdf' : 'fa-file-image'}" style="color: ${isPdf ? '#dc2626' : '#10b981'};"></i>
                    <span class="doc-type">Additional File</span>
                    <span class="doc-filename">(${file.name})</span>
                    ${!isPdf ? '<span class="merge-convert-badge">📷 Convert to PDF</span>' : ''}
                </label>
                <span class="merge-order-badge" id="order_${itemId}" style="display: ${isChecked ? 'inline-block' : 'none'};">${isChecked ? '#' + (mergeSelectionOrder.indexOf(itemId) + 1) : ''}</span>
                <button class="remove-item-btn" onclick="window.removeAdditionalFile('${itemId}')" title="Remove">
                    <i class="fas fa-times-circle"></i>
                </button>
            </div>
        `;
    });

    combinedHtml += '</div>';

    mergeList.innerHTML = combinedHtml;

    // Update select all checkbox
    const selectAllCheckbox = document.getElementById('selectAllDocuments');
    if (selectAllCheckbox) {
        const totalItems = originalDocs.length + additionalFiles.length;
        const checkedItems = mergeSelectionOrder.length;
        selectAllCheckbox.checked = checkedItems === totalItems && totalItems > 0;
        selectAllCheckbox.indeterminate = checkedItems > 0 && checkedItems < totalItems;
    }
}

// ==================== MERGE DOCUMENTS ====================
async function handleMergeDocuments() {
    const originalDocs = currentMergeSubmission?.documents || [];

    if (mergeSelectionOrder.length === 0) {
        showNotification('Please select documents to merge', 'error');
        return;
    }

    mergeInProgress = true;
    if (mergeDocumentsBtn) {
        mergeDocumentsBtn.disabled = true;
        mergeDocumentsBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Merging...';
    }

    if (mergeDownloadSection) {
        mergeDownloadSection.style.display = 'none';
    }

    try {
        const { PDFDocument } = PDFLib;
        const mergedPdf = await PDFDocument.create();

        const totalDocs = mergeSelectionOrder.length;
        let processedCount = 0;
        let successCount = 0;

        for (const itemId of mergeSelectionOrder) {
            processedCount++;

            try {
                if (itemId.startsWith('orig_')) {
                    const index = parseInt(itemId.replace('orig_', ''));
                    const doc = originalDocs[index];
                    if (!doc || !doc.fileUrl) {
                        showNotification(`⚠️ Skipping document ${processedCount}: Invalid document`, 'warning');
                        continue;
                    }

                    showNotification(`📄 Processing ${processedCount}/${totalDocs}: ${DOCUMENT_TYPES[doc.documentType] || 'Document'}...`, 'info');

                    const response = await fetchWithCorsFallback(doc.fileUrl);

                    if (!response.ok) {
                        showNotification(`⚠️ Failed to fetch document ${processedCount} (${response.status}), skipping...`, 'warning');
                        continue;
                    }

                    const blob = await response.blob();
                    const fileName = doc.name || 'file.pdf';

                    if (blob.size === 0) {
                        showNotification(`⚠️ Document ${processedCount} is empty, skipping...`, 'warning');
                        continue;
                    }

                    const arrayBuffer = await blob.arrayBuffer();

                    // Check if it's a PDF
                    const isPdf = blob.type.includes('pdf') ||
                                 fileName.toLowerCase().endsWith('.pdf') ||
                                 (arrayBuffer.byteLength > 4 &&
                                  new Uint8Array(arrayBuffer.slice(0, 4))[0] === 0x25);

                    if (isPdf) {
                        try {
                            const pdf = await PDFDocument.load(arrayBuffer, {
                                ignoreEncryption: true,
                                throwOnInvalidObject: false
                            });
                            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                            copiedPages.forEach(page => mergedPdf.addPage(page));
                            successCount++;
                        } catch (pdfError) {
                            console.error('PDF load error:', pdfError);
                            // Try as image
                            if (isImageFile(fileName)) {
                                try {
                                    const pdfBytes = await imageToPdf(blob, fileName);
                                    const pdf = await PDFDocument.load(pdfBytes);
                                    const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                                    copiedPages.forEach(page => mergedPdf.addPage(page));
                                    successCount++;
                                } catch (convertError) {
                                    throw new Error(`Failed to convert: ${convertError.message}`);
                                }
                            } else {
                                throw new Error(`Invalid PDF file: ${doc.name || 'document'}`);
                            }
                        }
                    } else if (isImageFile(fileName)) {
                        try {
                            const pdfBytes = await imageToPdf(blob, fileName);
                            const pdf = await PDFDocument.load(pdfBytes);
                            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                            copiedPages.forEach(page => mergedPdf.addPage(page));
                            successCount++;
                        } catch (convertError) {
                            throw new Error(`Failed to convert image: ${convertError.message}`);
                        }
                    } else {
                        showNotification(`⚠️ Unsupported file format: ${fileName}, skipping...`, 'warning');
                    }
                } else {
                    const fileObj = additionalFiles.find(f => f.id === itemId);
                    if (!fileObj) {
                        showNotification(`⚠️ Skipping file ${processedCount}: File not found`, 'warning');
                        continue;
                    }

                    const file = fileObj.file;
                    showNotification(`📄 Processing ${processedCount}/${totalDocs}: ${file.name}...`, 'info');

                    const arrayBuffer = await file.arrayBuffer();
                    const fileName = file.name.toLowerCase();

                    const isPdf = file.type.includes('pdf') ||
                                 fileName.endsWith('.pdf') ||
                                 (arrayBuffer.byteLength > 4 &&
                                  new Uint8Array(arrayBuffer.slice(0, 4))[0] === 0x25);

                    if (isPdf) {
                        try {
                            const pdf = await PDFDocument.load(arrayBuffer, {
                                ignoreEncryption: true,
                                throwOnInvalidObject: false
                            });
                            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                            copiedPages.forEach(page => mergedPdf.addPage(page));
                            successCount++;
                        } catch (pdfError) {
                            throw new Error(`Invalid PDF file: ${file.name}`);
                        }
                    } else if (fileName.endsWith('.jpg') || fileName.endsWith('.jpeg') ||
                             fileName.endsWith('.png') || fileName.endsWith('.gif') ||
                             fileName.endsWith('.bmp') || fileName.endsWith('.webp')) {
                        try {
                            const pdfBytes = await imageToPdf(file, file.name);
                            const pdf = await PDFDocument.load(pdfBytes);
                            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                            copiedPages.forEach(page => mergedPdf.addPage(page));
                            successCount++;
                        } catch (convertError) {
                            throw new Error(`Failed to convert image: ${convertError.message}`);
                        }
                    } else {
                        showNotification(`⚠️ Unsupported file format: ${file.name}, skipping...`, 'warning');
                    }
                }
            } catch (itemError) {
                console.error('Error processing item:', itemError);
                showNotification(`⚠️ Error processing item ${processedCount}: ${itemError.message}`, 'warning');
            }
        }

        if (mergedPdf.getPageCount() === 0) {
            throw new Error('No valid pages were merged. Please check your files.');
        }

        showNotification(`✅ Successfully merged ${successCount} out of ${totalDocs} documents (${mergedPdf.getPageCount()} pages)`, 'success');

        const mergedPdfBytes = await mergedPdf.save();
        mergedPdfBlob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
        if (mergedPdfUrl) {
            URL.revokeObjectURL(mergedPdfUrl);
        }
        mergedPdfUrl = URL.createObjectURL(mergedPdfBlob);

        if (mergeDownloadSection) {
            mergeDownloadSection.style.display = 'block';
            mergeDownloadSection.innerHTML = `
                <div class="merge-success-message">
                    <i class="fas fa-check-circle"></i>
                    <span>PDF merged successfully! (${mergedPdf.getPageCount()} pages from ${successCount} documents)</span>
                </div>
                <div class="merge-action-buttons">
                    <button class="action-btn view-merged-btn" id="viewMergedPdf">
                        <i class="fas fa-eye"></i> View Merged PDF
                    </button>
                    <button class="action-btn save-merged-btn" id="saveMergedPdf">
                        <i class="fas fa-save"></i> Save PDF
                    </button>
                </div>
            `;

            document.getElementById('viewMergedPdf')?.addEventListener('click', viewMergedPDF);
            document.getElementById('saveMergedPdf')?.addEventListener('click', saveMergedPDF);
        }

    } catch (error) {
        console.error('Merge error:', error);
        showNotification('Failed to merge: ' + error.message, 'error');
    } finally {
        mergeInProgress = false;
        if (mergeDocumentsBtn) {
            mergeDocumentsBtn.disabled = false;
            mergeDocumentsBtn.innerHTML = '<i class="fas fa-compress-alt"></i> Merge Selected';
        }
    }
}

// ==================== SAVE MERGED PDF ====================
async function saveMergedPDF() {
    if (!mergedPdfBlob) {
        showNotification('No merged PDF available', 'error');
        return;
    }

    const customerName = currentMergeSubmission?.customerName || 'Customer';
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const fileName = `${customerName.replace(/[^a-zA-Z0-9]/g, '_')}_Merged_${timestamp}.pdf`;

    await saveFileWithLocationPicker(mergedPdfBlob, fileName);
}

function viewMergedPDF() {
    if (!mergedPdfBlob || !mergedPdfUrl) {
        showNotification('No merged PDF available', 'error');
        return;
    }

    const customerName = currentMergeSubmission?.customerName || 'Customer';
    viewerFileName.textContent = `${customerName} - Merged Documents`;
    documentViewer.src = mergedPdfUrl;

    if (viewerSaveBtn) {
        viewerSaveBtn.style.display = 'none';
    }

    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'block';
    }

    // Clear navigation for merged view
    if (viewerNav) {
        viewerNav.innerHTML = '';
    }

    viewerModal.classList.add('active');
    currentZoom = 1.0;
    applyZoom();

    // Keep merge modal open when viewing merged PDF
    // closeMergeModalFunc();
}

function closeMergeModalFunc() {
    mergeModal.classList.remove('active');
    currentMergeSubmission = null;
    mergeSelectionOrder = [];
    additionalFiles = [];
    additionalFileCounter = 0;
    mergedPdfBlob = null;
    if (mergedPdfUrl) {
        URL.revokeObjectURL(mergedPdfUrl);
        mergedPdfUrl = null;
    }
    if (mergeDownloadSection) {
        mergeDownloadSection.style.display = 'none';
    }
    updateSelectedCount();
}

// ==================== TAB SWITCHING ====================
window.switchTab = (tabId) => {
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    event.currentTarget.classList.add('active');

    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.getElementById(tabId + 'Tab').classList.add('active');

    const titles = {
        'pending': 'Pending Documents Review',
        'approved': 'Approved Documents',
        'rejected': 'Rejected Documents'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId];
};

// ==================== MODAL UTILITIES ====================
function closeModal(modal) {
    modal.classList.remove('active');
    if (modal === commentModal) {
        currentSubmissionId = null;
        commentText.value = '';
    }
}

// ==================== NOTIFICATION ====================
function showNotification(message, type = 'info') {
    if (!notification) return;

    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    notification.classList.add('show');

    setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => {
            notification.style.display = 'none';
        }, 300);
    }, 3000);
}

// ==================== CUSTOMER DETAILS MODAL ====================
function formatCurrency(value) {
    const num = Number(value || 0);
    if (isNaN(num)) {
        return '₦0.00';
    } else {
        return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
}

window.openCustomerDetails = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    const details = sub.customerDetails || {};
    const doc = document;

    // Populate personal information
    doc.getElementById('customerDetailsTitle').textContent = `Details - ${sub.customerName}`;
    doc.getElementById('detailCustomerName').textContent = sub.customerName || '-';
    doc.getElementById('detailDOB').textContent = details.dob || '-';
    doc.getElementById('detailNIN').textContent = details.nin || '-';
    doc.getElementById('detailPFA').textContent = details.pfa || '-';
    doc.getElementById('detailAccountNo').textContent = details.accountNo || '-';
    doc.getElementById('detailEmail').textContent = details.email || '-';
    doc.getElementById('detailPhone').textContent = details.phone || '-';
    doc.getElementById('detailAddress').textContent = details.address || '-';
    doc.getElementById('detailEmployer').textContent = details.employer || '-';
    doc.getElementById('detailOriginatingTP').textContent = details.originatingTP || '-';
    doc.getElementById('detailPenNo').textContent = details.penNo || '-';

    // Populate RSA information
    doc.getElementById('detailRSABalance').textContent = details.rsaBalance ? formatCurrency(details.rsaBalance) : '-';
    doc.getElementById('detailRSADate').textContent = details.rsaStatementDate || '-';
    doc.getElementById('detailRSA25').textContent = details.rsa25 ? formatCurrency(details.rsa25) : '-';

    // Property / loan
    doc.getElementById('detailPropertyType').textContent = details.propertyType || '-';
    doc.getElementById('detailTenor').textContent = details.tenor ? `${details.tenor}` : '-';
    doc.getElementById('detailPropertyValue').textContent = details.propertyValue ? formatCurrency(details.propertyValue) : '-';
    doc.getElementById('detailFacilityFee').textContent = details.facilityFee ? formatCurrency(details.facilityFee) : '-';
    doc.getElementById('detailLoanAmount').textContent = details.loanAmount ? formatCurrency(details.loanAmount) : '-';

    // Populate submission information
    doc.getElementById('detailUploadedBy').textContent = sub.uploadedBy || '-';
    doc.getElementById('detailUploadedAt').textContent = formatTimestamp(sub.uploadedAt);
    doc.getElementById('detailStatus').textContent = (sub.status || '-').toUpperCase();
    doc.getElementById('detailReviewedBy').textContent = sub.reviewedBy || '-';

    // Populate documents list
    const docsList = sub.documents || [];

    const docListHtml = docsList.length === 0
        ? '<p style="color: #999; text-align: center; padding: 20px;">No documents submitted</p>'
        : `<div style="display: grid; gap: 10px;">
            ${docsList.map((doc, idx) => `
                <div style="background: white; padding: 12px; border: 1px solid #e5e7eb; border-radius: 6px; display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <div style="font-weight: 600;">${DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document'}</div>
                        <div style="font-size: 12px; color: #666;">${doc.name}</div>
                    </div>
                    <div style="display: flex; gap: 8px;">
                        <button class="action-btn" onclick="window.viewDocument('${doc.fileUrl}', '${doc.name}')" title="View">
                            <i class="fas fa-eye"></i>
                        </button>
                    </div>
                </div>
            `).join('')}
        </div>`;

    doc.getElementById('customerDocumentsList').innerHTML = docListHtml;

    // Show modal
    const modal = doc.getElementById('customerDetailsModal');
    if (modal) modal.classList.add('active');
};

function closeCustomerDetailsModal() {
    const modal = document.getElementById('customerDetailsModal');
    if (modal) modal.classList.remove('active');
}

// Add event listeners for close buttons
document.getElementById('closeCustomerDetails')?.addEventListener('click', closeCustomerDetailsModal);
document.getElementById('closeCustomerDetailsFooter')?.addEventListener('click', closeCustomerDetailsModal);

// Close modal when clicking outside
window.addEventListener('click', (e) => {
    const customerDetailsModal = document.getElementById('customerDetailsModal');
    if (e.target === customerDetailsModal) closeCustomerDetailsModal();
});

// ==================== SIGN OUT ====================
window.signOutUser = () => {
    window.location.href = 'index.html';
};

// Make functions global
window.viewSubmission = viewSubmission;
window.openReviewModal = openReviewModal;
window.viewDocument = viewDocument;
window.switchTab = switchTab;
window.openMergeModal = openMergeModal;
window.handleMergeCheckbox = handleMergeCheckbox;
window.selectAllDocuments = selectAllDocuments;
window.downloadAll = downloadAll;
window.viewMergedPDF = viewMergedPDF;
window.saveMergedPDF = saveMergedPDF;
window.clearAdditionalFiles = clearAdditionalFiles;
window.removeAdditionalFile = removeAdditionalFile;
