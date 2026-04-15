import { auth, db } from './firebase-config.js';
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    addDoc,
    doc,
    query,
    where,
    orderBy,
    onSnapshot,
    getDocs,
    updateDoc,
    runTransaction,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { queueUploaderFinalSubmissionEmail } from './email-alerts.js';
import { notifyStatusChangePush } from './status-push.js';

// ==================== DOCUMENT TYPES MAPPING ====================
const DOCUMENT_TYPES = {
    'birth_certificate': 'Birth Certificate / Age Declaration',
    'nin': 'National Identification Number (NIN)',
    'bvn': 'BVN',
    'pay_slips': '3 Months Pay Slip',
    'offer_letter': 'Offer of Employment Letter',
    'intro_letter': 'Introduction Letter',
    'request_letter': 'Request Letter',
    'rsa_statement': 'RSA Statement',
    'pfa_form': 'PFA Application Form',
    'consent_letter': 'Consent Letter',
    'indemnity_form': 'Indemnity Form',
    'mortgage_loan_application': 'Mortgage Loan Application Form',
    'allocation_last_page': 'Allocation Last Page',
    'offer_letter_last_page': 'Offer Letter Last Page',
    'pmi_soa': 'PMI SOA',
    'benefit_application_form': 'Benefit Application Form',
    'data_recapture': 'Data Recapture'
};

let currentUser = null;
let currentRsaProfileData = null;
let currentTab = 'approved';
let allSubmissions = [];
const userFullNameCache = new Map();
let unsubscribeQueue = null;
let queueLoadSeq = 0;

const IDLE_TIMEOUT_MS = 30 * 60 * 1000;
let idleLastActivity = Date.now();
let idleIntervalHandle = null;

const rsaTableBody = document.getElementById('rsaTableBody');
const rsaSearch = document.getElementById('rsaSearch');
const rsaStartDate = document.getElementById('rsaStartDate');
const rsaEndDate = document.getElementById('rsaEndDate');
const rsaFilterBtn = document.getElementById('rsaFilterBtn');
const pageTitle = document.getElementById('pageTitle');
const approvedCountBadge = document.getElementById('approvedCount');
const userNameEl = document.getElementById('userName');
const userRoleEl = document.getElementById('userRole');
const finalSubmitBtn = document.getElementById('finalSubmitBtn');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');

let currentDetailsSubmissionId = null;

function renderProfileTab() {
    if (!profileNameEl && !profileEmailEl && !profileRoleEl && !profileStatusEl) return;
    const fullName = currentRsaProfileData?.fullName || currentRsaProfileData?.displayName || currentUser?.displayName || currentUser?.email || 'N/A';
    const registeredAt = currentRsaProfileData?.createdAt ? formatTimestamp(currentRsaProfileData.createdAt) : '-';
    const email = currentRsaProfileData?.email || currentUser?.email || 'N/A';
    const whatsapp = currentRsaProfileData?.whatsappNumber || currentRsaProfileData?.phone || '-';
    const location = currentRsaProfileData?.location || '-';
    const role = String(currentRsaProfileData?.role || 'rsa');
    const status = String(currentRsaProfileData?.status || 'active');
    if (profileNameEl) profileNameEl.textContent = fullName;
    if (profileRegisteredAtEl) profileRegisteredAtEl.textContent = registeredAt;
    if (profileEmailEl) profileEmailEl.textContent = email;
    if (profileWhatsappEl) profileWhatsappEl.textContent = whatsapp;
    if (profileLocationEl) profileLocationEl.textContent = location;
    if (profileRoleEl) profileRoleEl.textContent = role.toUpperCase();
    if (profileStatusEl) profileStatusEl.textContent = status.charAt(0).toUpperCase() + status.slice(1);
}

// format time helper
function formatTimestamp(ts) {
    if (!ts) return 'N/A';
    try {
        const d = ts.toDate ? ts.toDate() : new Date(ts);
        return d.toLocaleString('en-NG', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });
    } catch(e){return 'N/A';}
}

function normalizeWhatsAppPhone(raw) {
    const digits = String(raw || '').replace(/\D/g, '');
    if (!digits) return '';
    if (digits.startsWith('0') && digits.length === 11) return `234${digits.slice(1)}`;
    if (digits.length === 10) return `234${digits}`;
    if (digits.startsWith('234')) return digits;
    return digits;
}

function setWhatsAppContact(containerEl, rawPhone) {
    if (!containerEl) return;
    const display = String(rawPhone || '').trim();
    const normalized = normalizeWhatsAppPhone(display);
    if (!normalized) {
        containerEl.textContent = '-';
        return;
    }
    containerEl.innerHTML = `<a href="https://wa.me/${normalized}" target="_blank" rel="noopener noreferrer">${display}</a>`;
}

function getApprovedTimestamp(sub) {
    // Reviewer approval time is typically stored as reviewedAt/approvedAt.
    return sub?.reviewedAt || sub?.approvedAt || sub?.statusUpdatedAt || sub?.updatedAt || null;
}

function normalizeEmail(email) {
    return String(email || '').trim().toLowerCase();
}

function routingRuleDocId(uploaderEmail) {
    return encodeURIComponent(normalizeEmail(uploaderEmail));
}

async function getUploaderRoutingRule(uploaderEmail) {
    const normalizedUploader = normalizeEmail(uploaderEmail);
    if (!normalizedUploader) return null;
    try {
        const snap = await getDoc(doc(db, 'uploaderRoutingRules', routingRuleDocId(normalizedUploader)));
        if (!snap.exists()) return null;
        const data = snap.data() || {};
        if (data.enabled === false) return null;
        return {
            uploaderEmail: normalizedUploader,
            reviewerEmail: normalizeEmail(data.reviewerEmail),
            rsaEmail: normalizeEmail(data.rsaEmail),
            paymentEmail: normalizeEmail(data.paymentEmail)
        };
    } catch (_) {
        return null;
    }
}

async function isActivePaymentUser(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return false;
    try {
        const userQuery = query(collection(db, 'users'), where('email', '==', normalized));
        const snap = await getDocs(userQuery);
        if (snap.empty) return false;
        const data = snap.docs[0].data() || {};
        const role = String(data.role || '').toLowerCase();
        const status = String(data.status || 'active').toLowerCase();
        return role === 'payment' && status !== 'deactivated';
    } catch (_) {
        return false;
    }
}

async function getActivePaymentUsers() {
    const paymentQuery = query(collection(db, 'users'), where('role', '==', 'payment'));
    const snap = await getDocs(paymentQuery);
    return snap.docs
        .map((d) => d.data() || {})
        .map((u) => normalizeEmail(u.email))
        .filter(Boolean)
        .filter((email, idx, arr) => arr.indexOf(email) === idx)
        .sort();
}

async function assignPaymentRoundRobin(submissionId) {
    const submissionRef = doc(db, 'submissions', submissionId);
    let uploaderEmail = '';
    try {
        const subSnap = await getDoc(submissionRef);
        if (subSnap.exists()) uploaderEmail = normalizeEmail(subSnap.data()?.uploadedBy);
    } catch (_) { }

    const routingRule = await getUploaderRoutingRule(uploaderEmail);
    const mappedPayment = routingRule?.paymentEmail || '';
    if (mappedPayment && await isActivePaymentUser(mappedPayment)) {
        await updateDoc(submissionRef, {
            assignedToPayment: mappedPayment,
            paymentAssignedAt: serverTimestamp(),
            paymentAssignmentMethod: 'uploader_routing'
        });
        return mappedPayment;
    }

    const paymentUsers = await getActivePaymentUsers();

    if (!paymentUsers.length) {
        await updateDoc(submissionRef, {
            assignedToPayment: '',
            paymentAssignedAt: serverTimestamp(),
            paymentAssignmentMethod: 'unassigned'
        });
        return '';
    }

    const counterRef = doc(db, 'counters', 'roundRobinPayment');
    let assignedPayment = '';
    const today = new Date().toISOString().slice(0, 10);

    try {
        await runTransaction(db, async (tx) => {
            const counterSnap = await tx.get(counterRef);
            let lastIndex = -1;
            let lastDate = '';

            if (counterSnap.exists()) {
                const data = counterSnap.data() || {};
                lastIndex = typeof data.lastIndex === 'number' ? data.lastIndex : -1;
                lastDate = data.lastDate || '';
            }

            if (lastDate !== today) lastIndex = -1;
            const nextIndex = (lastIndex + 1) % paymentUsers.length;
            assignedPayment = paymentUsers[nextIndex];

            tx.set(counterRef, {
                lastIndex: nextIndex,
                lastDate: today,
                lastAssignedTo: assignedPayment,
                updatedAt: serverTimestamp()
            }, { merge: true });

            tx.update(submissionRef, {
                assignedToPayment: assignedPayment,
                paymentAssignedAt: serverTimestamp(),
                paymentAssignmentMethod: 'round_robin'
            });
        });
    } catch (_) {
        assignedPayment = paymentUsers[0] || '';
        if (assignedPayment) {
            await updateDoc(submissionRef, {
                assignedToPayment: assignedPayment,
                paymentAssignedAt: serverTimestamp(),
                paymentAssignmentMethod: 'round_robin_fallback'
            });
        }
    }

    return assignedPayment;
}

async function getUserFullName(email) {
    const normalizedEmail = String(email || '').trim().toLowerCase();
    if (!normalizedEmail) return '';
    if (userFullNameCache.has(normalizedEmail)) return userFullNameCache.get(normalizedEmail);

    try {
        const userQuery = query(collection(db, 'users'), where('email', '==', normalizedEmail));
        const userSnapshot = await getDocs(userQuery);
        if (!userSnapshot.empty) {
            const userData = userSnapshot.docs[0].data();
            const fullName = userData.fullName || userData.displayName || normalizedEmail.split('@')[0];
            userFullNameCache.set(normalizedEmail, fullName);
            return fullName;
        }
    } catch (error) {
        console.warn('Could not load user full name:', normalizedEmail, error);
    }

    const fallbackName = normalizedEmail.split('@')[0];
    userFullNameCache.set(normalizedEmail, fallbackName);
    return fallbackName;
}

async function loadCurrentRsaProfile(user) {
    if (!user) return false;

    const email = String(user.email || '').trim().toLowerCase();
    let profileData = null;

    try {
        const userQuery = query(collection(db, 'users'), where('email', '==', email));
        const userSnapshot = await getDocs(userQuery);
        if (!userSnapshot.empty) {
            profileData = userSnapshot.docs[0].data();
            currentRsaProfileData = profileData;
        }
    } catch (error) {
        console.error('Failed to fetch RSA user profile:', error);
    }

    const displayName = profileData?.fullName || profileData?.displayName || user.displayName || email.split('@')[0] || 'RSA User';
    if (userNameEl) userNameEl.textContent = displayName;
    if (userRoleEl) userRoleEl.textContent = 'RSA';
    if (!currentRsaProfileData) {
        currentRsaProfileData = { email, fullName: displayName, role: 'rsa', status: 'active' };
    }
    renderProfileTab();

    const role = String(profileData?.role || '').toLowerCase();
    if (role && role !== 'rsa') {
        showNotification('Access denied. RSA privileges required.', 'error');
        setTimeout(() => { window.location.href = 'index.html'; }, 1500);
        return false;
    }
    return true;
}

// ------- utility helpers (borrowed from viewer dashboard) -------
async function fetchWithCorsFallback(url) {
    const cleanUrl = url?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) throw new Error('Invalid URL');

    const response = await fetch(cleanUrl, { mode: 'cors', credentials: 'omit' });
    if (!response.ok) throw new Error(`Document fetch failed: ${response.status}`);
    return response;
}

// PDF/image helpers
function isPdfFile(fileUrl, fileName) {
    const lowerName = (fileName||'').toLowerCase();
    const lowerUrl = (fileUrl||'').toLowerCase();
    return lowerName.endsWith('.pdf') || lowerUrl.includes('.pdf') || lowerUrl.includes('application/pdf');
}
function isImageFile(fileName) {
    const lowerName = fileName.toLowerCase();
    return lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg') || lowerName.endsWith('.png') || lowerName.endsWith('.gif') || lowerName.endsWith('.bmp') || lowerName.endsWith('.webp');
}
async function imageToPdf(imageBlob, imageName) {
    try {
        const { PDFDocument } = PDFLib;
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([595, 842]);
        const imageBytes = await imageBlob.arrayBuffer();
        let image;
        const lowerName = imageName.toLowerCase();
        const imageType = imageBlob.type || '';
        try {
            if (lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg') || imageType.includes('jpeg')) {
                image = await pdfDoc.embedJpg(imageBytes);
            } else if (lowerName.endsWith('.png') || imageType.includes('png')) {
                image = await pdfDoc.embedPng(imageBytes);
            } else {
                const arr = new Uint8Array(imageBytes.slice(0,4));
                if (arr[0]===0xFF && arr[1]===0xD8) {
                    image = await pdfDoc.embedJpg(imageBytes);
                } else if (arr[0]===0x89 && arr[1]===0x50 && arr[2]===0x4E && arr[3]===0x47) {
                    image = await pdfDoc.embedPng(imageBytes);
                } else {
                    throw new Error(`Unsupported image format: ${imageName}`);
                }
            }
        } catch (embedError) {
            throw embedError;
        }
        const imgDims = image.scale(0.5);
        page.drawImage(image, { x: 50, y: page.getHeight() - imgDims.height - 50, width: imgDims.width, height: imgDims.height });
        const pdfBytes = await pdfDoc.save();
        return pdfBytes;
    } catch (e) {
        console.error('imageToPdf error',e);
        throw e;
    }
}

function showNotification(message, type='info') {
    const notification = document.getElementById('notification');
    if (!notification) return;
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    setTimeout(()=>{ notification.style.display='none'; }, 3000);
}

function showLoader(msg) {
    const loader = document.getElementById('globalLoader');
    const text = document.getElementById('loaderText');
    if (loader && text) {
        text.textContent = msg || 'Processing...';
        loader.style.display = 'flex';
        setTimeout(()=>loader.classList.add('active'),10);
    }
}
function hideLoader() {
    const loader = document.getElementById('globalLoader');
    if (loader) {
        loader.classList.remove('active');
        setTimeout(()=>loader.style.display='none',300);
    }
}

async function saveBlobToFolderPicker(blob, defaultFileName, customerName='Customer') {
    if (!('showDirectoryPicker' in window)) {
        showNotification('Folder picker not supported. Falling back to save dialog...', 'info');
        return saveFileWithLocationPicker(blob, defaultFileName);
    }
    try {
        showNotification('📁 Please select a destination folder...', 'info');
        const dirHandle = await window.showDirectoryPicker({ mode:'readwrite', startIn:'downloads' });
        const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s_-]/g,'_').trim() || 'Customer';
        const customerFolder = await dirHandle.getDirectoryHandle(safeCustomerName, { create:true });
        const fileHandle = await customerFolder.getFileHandle(defaultFileName,{ create:true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        showNotification(`✅ Saved to ${safeCustomerName}/${defaultFileName}`,'success');
        return true;
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Save cancelled','info');
        } else {
            console.error('Save error (folder picker):',error);
            showNotification('Save failed: '+error.message,'error');
            await saveFileWithLocationPicker(blob, defaultFileName);
        }
        return false;
    }
}

function downloadBlobAsFile(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showNotification('✅ Download started','success');
}

// ==================== MERGE FUNCTIONS COPY START ====================
let currentMergeSubmission = null;
let mergeSelectionOrder = [];
let additionalFiles = [];
let additionalFileCounter = 0;
let mergedPdfBlob = null;
let mergedPdfUrl = null;
let mergeInProgress = false;

window.openMergeModal = (submissionId) => {
    const sub = allSubmissions.find(s=>s.id===submissionId);
    if (!sub) return;
    currentMergeSubmission = sub;
    mergeSelectionOrder = [];
    additionalFiles = [];
    additionalFileCounter = 0;
    mergedPdfBlob = null;
    if (mergedPdfUrl) { URL.revokeObjectURL(mergedPdfUrl); mergedPdfUrl = null; }
    const mergeDownloadSection = document.getElementById('mergeDownloadSection');
    if (mergeDownloadSection) mergeDownloadSection.style.display='none';
    document.getElementById('mergeCustomerName').textContent = sub.customerName;
    if (!sub.documents || sub.documents.length===0) {
        showNotification('No documents available to merge','error');
        return;
    }
    refreshMergeList();
    const previewContainer = document.getElementById('additionalFilesPreview');
    if (previewContainer) previewContainer.innerHTML='';
    const fileInput = document.getElementById('additionalFileInput');
    if (fileInput) fileInput.value='';
    setTimeout(setupAdditionalFileUpload,100);
    document.getElementById('mergeModal').classList.add('active');
    updateSelectedCount();
};

window.handleMergeCheckbox = (checkbox,itemId) => {
    if (checkbox.checked) {
        if (!mergeSelectionOrder.includes(itemId)) mergeSelectionOrder.push(itemId);
    } else {
        mergeSelectionOrder = mergeSelectionOrder.filter(id=>id!==itemId);
    }
    renumberBadges();
    updateSelectedCount();
    const selectAllCheckbox = document.getElementById('selectAllDocuments');
    if (selectAllCheckbox) {
        const totalItems = document.querySelectorAll('.merge-item input[type="checkbox"]').length;
        const checkedItems = mergeSelectionOrder.length;
        selectAllCheckbox.checked = checkedItems === totalItems && totalItems > 0;
        selectAllCheckbox.indeterminate = checkedItems>0 && checkedItems<totalItems;
    }
};

window.selectAllDocuments = (checkbox) => {
    const checkboxes = document.querySelectorAll('.merge-item input[type="checkbox"]');
    mergeSelectionOrder = [];
    if (checkbox.checked) {
        checkboxes.forEach(cb=>{ cb.checked=true; mergeSelectionOrder.push(cb.value); });
    } else {
        checkboxes.forEach(cb=>{ cb.checked=false; });
    }
    renumberBadges();
    updateSelectedCount();
};

function updateSelectedCount() {
    const selectedCount = document.getElementById('selectedCount');
    if (selectedCount) selectedCount.textContent = mergeSelectionOrder.length;
}

function renumberBadges() {
    mergeSelectionOrder.forEach((itemId,order)=>{
        const badge = document.getElementById(`order_${itemId}`);
        if (badge) { badge.textContent = `#${order+1}`; badge.style.display='inline-block'; }
    });
    document.querySelectorAll('.merge-order-badge').forEach(badge=>{
        if (!mergeSelectionOrder.includes(badge.id.replace('order_',''))) {
            badge.textContent=''; badge.style.display='none';
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
    originalDocs.forEach((doc,idx)=>{
        const itemId = `orig_${idx}`;
        const isChecked = mergeSelectionOrder.includes(itemId);
        const isPdf = isPdfFile(doc.fileUrl, doc.name);
        combinedHtml += `
            <div class="merge-item" data-url="${doc.fileUrl}" data-name="${doc.name}" id="item_${itemId}">
                <input type="checkbox" value="${itemId}" ${isChecked?'checked':''} onchange="window.handleMergeCheckbox(this,'${itemId}')">
                <span class="merge-order-badge" id="order_${itemId}" style="display: ${isChecked?'inline-block':'none'};">${isChecked?`#${mergeSelectionOrder.indexOf(itemId)+1}`:''}</span>
                <div class="merge-thumb"></div>
                <div class="merge-info">
                    <div>${DOCUMENT_TYPES[doc.documentType]||doc.documentType||'Document'}</div>
                    <div class="doc-filename">${doc.name}</div>
                </div>
            </div>
        `;
    });
    additionalFiles.forEach(f=>{
        const itemId = f.id;
        const isChecked = mergeSelectionOrder.includes(itemId);
        const isPdf = isPdfFile('', f.name);
        combinedHtml += `
            <div class="merge-item additional" data-url="" data-name="${f.name}" id="item_${itemId}">
                <input type="checkbox" value="${itemId}" ${isChecked?'checked':''} onchange="window.handleMergeCheckbox(this,'${itemId}')">
                <span class="merge-order-badge" id="order_${itemId}" style="display: ${isChecked?'inline-block':'none'};">${isChecked?`#${mergeSelectionOrder.indexOf(itemId)+1}`:''}</span>
                <div class="merge-thumb"></div>
                <div class="merge-info">
                    <div>${f.name}</div>
                    <div class="doc-filename">(additional)</div>
                </div>
                <button class="remove-item-btn" onclick="window.removeAdditionalFile('${itemId}')" title="Remove">
                    <i class="fas fa-times-circle"></i>
                </button>
            </div>
        `;
    });
    combinedHtml += '</div>';
    const mergeListEl = document.getElementById('mergeList');
    if (mergeListEl) mergeListEl.innerHTML = combinedHtml;
    const selectAllCheckbox = document.getElementById('selectAllDocuments');
    if (selectAllCheckbox) {
        const totalItems = originalDocs.length + additionalFiles.length;
        const checkedItems = mergeSelectionOrder.length;
        selectAllCheckbox.checked = checkedItems===totalItems && totalItems>0;
        selectAllCheckbox.indeterminate = checkedItems>0 && checkedItems<totalItems;
    }
    renderMergeThumbnails();
}

async function renderMergeThumbnails() {
    // Thumbnails removed (pdfThumbnail was not included in RSA dashboard build).
    // Keep placeholder empty to avoid runtime errors.
    return;
}

function setupAdditionalFileUpload() {
    const uploadArea = document.getElementById('additionalUploadArea');
    const fileInput = document.getElementById('additionalFileInput');
    if (!uploadArea || !fileInput) return;
    const newUploadArea = uploadArea.cloneNode(true);
    uploadArea.parentNode.replaceChild(newUploadArea, uploadArea);
    const newFileInput = fileInput.cloneNode(true);
    fileInput.parentNode.replaceChild(newFileInput, fileInput);
    const updatedUploadArea = document.getElementById('additionalUploadArea');
    const updatedFileInput = document.getElementById('additionalFileInput');
    updatedUploadArea.addEventListener('click', () => { updatedFileInput.click(); });
    updatedUploadArea.addEventListener('dragover', (e)=>{
        e.preventDefault();
        updatedUploadArea.style.borderColor='var(--cm-primary)';
        updatedUploadArea.style.background='var(--cm-light)';
    });
    updatedUploadArea.addEventListener('dragleave', ()=>{
        updatedUploadArea.style.borderColor='#cbd5e1';
        updatedUploadArea.style.background='transparent';
    });
    updatedUploadArea.addEventListener('drop', (e)=>{
        e.preventDefault();
        updatedUploadArea.style.borderColor='#cbd5e1';
        updatedUploadArea.style.background='transparent';
        const files = Array.from(e.dataTransfer.files);
        handleAdditionalFiles(files);
    });
    updatedFileInput.addEventListener('change', (e)=>{
        const files = Array.from(e.target.files);
        handleAdditionalFiles(files);
    });
}

function handleAdditionalFiles(files) {
    if (files.length + additionalFiles.length > 10) {
        showNotification('Maximum 10 files total allowed','error');
        return;
    }
    const allowedTypes = ['.pdf','.jpg','.jpeg','.png','.gif','.bmp','.webp'];
    const validFiles = files.filter(file => {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        return allowedTypes.includes(ext);
    });
    if (validFiles.length !== files.length) {
        showNotification('Some files were skipped (only PDF and common image formats allowed)','warning');
    }
    validFiles.forEach(file=>{
        const fileId = `additional_${additionalFileCounter++}`;
        additionalFiles.push({ id:fileId, file:file, name:file.name, type:file.type, size:file.size, isAdditional:true });
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
    previewContainer.innerHTML = additionalFiles.map(f=>{
        return `<div class="file-preview" id="preview_${f.id}">
            <span>${f.name}</span>
            <button class="remove-file" onclick="window.removeAdditionalFile('${f.id}')"><i class="fas fa-times-circle"></i></button>
        </div>`;
    }).join('');
}

window.removeAdditionalFile = (id) => {
    additionalFiles = additionalFiles.filter(f=>f.id!==id);
    refreshMergeList();
    updateFilePreview();
    mergeSelectionOrder = mergeSelectionOrder.filter(i=>i!==id);
    updateSelectedCount();
};

async function handleMergeDocuments() {
    const originalDocs = currentMergeSubmission?.documents || [];
    if (mergeSelectionOrder.length === 0) {
        showNotification('Please select documents to merge','error');
        return;
    }
    mergeInProgress = true;
    const mergeDocumentsBtn = document.getElementById('mergeDocumentsBtn');
    if (mergeDocumentsBtn) {
        mergeDocumentsBtn.disabled = true;
        mergeDocumentsBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Merging...';
    }
    const mergeDownloadSection = document.getElementById('mergeDownloadSection');
    if (mergeDownloadSection) mergeDownloadSection.style.display='none';
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
                    const index = parseInt(itemId.replace('orig_',''));
                    const doc = originalDocs[index];
                    if (!doc || !doc.fileUrl) {
                        showNotification(`⚠️ Skipping document ${processedCount}: Invalid document`,'warning');
                        continue;
                    }
                    showNotification(`📄 Processing ${processedCount}/${totalDocs}: ${DOCUMENT_TYPES[doc.documentType]||'Document'}...`,'info');
                    let response;
                    try {
                        response = await fetchWithCorsFallback(doc.fileUrl);
                    } catch (fetchError) {
                        console.error('Fetch error:',fetchError);
                        showNotification(`⚠️ Could not fetch document ${processedCount}, skipping...`,'warning');
                        continue;
                    }
                    if (!response.ok) {
                        showNotification(`⚠️ Failed to fetch document ${processedCount} (${response.status}), skipping...`,'warning');
                        continue;
                    }
                    const blob = await response.blob();
                    const fileName = doc.name || 'file.pdf';
                    if (blob.size === 0) {
                        showNotification(`⚠️ Document ${processedCount} is empty, skipping...`,'warning');
                        continue;
                    }
                    const arrayBuffer = await blob.arrayBuffer();
                    const isPdf = blob.type.includes('pdf') || fileName.toLowerCase().endsWith('.pdf') || (arrayBuffer.byteLength>4 && new Uint8Array(arrayBuffer.slice(0,4))[0]===0x25);
                    if (isPdf) {
                        try {
                            const pdf = await PDFDocument.load(arrayBuffer,{ignoreEncryption:true,throwOnInvalidObject:false});
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (pdfError) {
                            console.error('PDF load error:',pdfError);
                            if (isImageFile(fileName)) {
                                showNotification(`🔄 Attempting to convert as image...`,'info');
                                try {
                                    const pdfBytes = await imageToPdf(blob,fileName);
                                    const pdf = await PDFDocument.load(pdfBytes);
                                    const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                                    copiedPages.forEach(page=>mergedPdf.addPage(page));
                                    successCount++;
                                } catch (convertError) {
                                    throw new Error(`Failed to convert: ${convertError.message}`);
                                }
                            } else {
                                throw new Error(`Invalid PDF file: ${doc.name||'document'}`);
                            }
                        }
                    } else if (isImageFile(fileName)) {
                        showNotification(`🔄 Converting image to PDF...`,'info');
                        try {
                            const pdfBytes = await imageToPdf(blob,fileName);
                            const pdf = await PDFDocument.load(pdfBytes);
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (convertError) {
                            throw new Error(`Failed to convert image: ${convertError.message}`);
                        }
                    } else {
                        showNotification(`⚠️ Unsupported file format: ${fileName}, skipping...`,'warning');
                    }
                } else {
                    const fileObj = additionalFiles.find(f=>f.id===itemId);
                    if (!fileObj) {
                        showNotification(`⚠️ Skipping file ${processedCount}: File not found`,'warning');
                        continue;
                    }
                    const file = fileObj.file;
                    showNotification(`📄 Processing ${processedCount}/${totalDocs}: ${file.name}...`,'info');
                    const arrayBuffer = await file.arrayBuffer();
                    const fileName = file.name.toLowerCase();
                    const isPdf = file.type.includes('pdf') || fileName.endsWith('.pdf') || (arrayBuffer.byteLength>4 && new Uint8Array(arrayBuffer.slice(0,4))[0]===0x25);
                    if (isPdf) {
                        try {
                            const pdf = await PDFDocument.load(arrayBuffer,{ignoreEncryption:true,throwOnInvalidObject:false});
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (pdfError) {
                            console.error('PDF load error:',pdfError);
                            throw new Error(`Invalid PDF file: ${file.name}`);
                        }
                    } else if (fileName.match(/\.(jpg|jpeg|png|gif|bmp|webp)$/)) {
                        showNotification(`🔄 Converting image to PDF: ${file.name}...`,'info');
                        try {
                            const pdfBytes = await imageToPdf(file,file.name);
                            const pdf = await PDFDocument.load(pdfBytes);
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (convertError) {
                            throw new Error(`Failed to convert image: ${convertError.message}`);
                        }
                    } else {
                        showNotification(`⚠️ Unsupported file format: ${file.name}, skipping...`,'warning');
                    }
                }
            } catch (itemError) {
                console.error('Error processing item:',itemError);
                showNotification(`⚠️ Error processing item ${processedCount}: ${itemError.message}`,'warning');
            }
        }
        if (mergedPdf.getPageCount() === 0) {
            throw new Error('No valid pages were merged. Please check your files.');
        }
        showNotification(`✅ Successfully merged ${successCount} out of ${totalDocs} documents (${mergedPdf.getPageCount()} pages)`,'success');
        showNotification('⏳ Finalizing merged document...','info');
        const mergedPdfBytes = await mergedPdf.save();
        mergedPdfBlob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
        if (mergedPdfUrl) URL.revokeObjectURL(mergedPdfUrl);
        mergedPdfUrl = URL.createObjectURL(mergedPdfBlob);
        if (mergeDownloadSection) {
            mergeDownloadSection.style.display='block';
            mergeDownloadSection.innerHTML = `
                <div class="merge-success-message">
                    <i class="fas fa-check-circle"></i>
                    <span>PDF merged successfully! (${mergedPdf.getPageCount()} pages from ${successCount} documents)</span>
                </div>
                <div class="merge-action-buttons">
                    <button class="action-btn view-merged-btn" id="viewMergedPdf">
                        <i class="fas fa-eye"></i> View Merged PDF
                    </button>
                    <button class="action-btn download-merged-btn" id="saveMergedPdf">
                        <i class="fas fa-save"></i> Save PDF
                    </button>
                </div>
            `;
            document.getElementById('viewMergedPdf')?.addEventListener('click', viewMergedPDF);
            document.getElementById('saveMergedPdf')?.addEventListener('click', saveMergedPDFWithPicker);
        }
    } catch (error) {
        console.error('Merge error:',error);
        showNotification('Failed to merge: ' + error.message,'error');
    } finally {
        mergeInProgress = false;
        if (mergeDocumentsBtn) {
            mergeDocumentsBtn.disabled = false;
            mergeDocumentsBtn.innerHTML = '<i class="fas fa-compress-alt"></i> Merge Selected';
        }
    }
}

async function saveMergedPDFWithPicker() {
    if (!mergedPdfBlob) {
        showNotification('No merged PDF available','error');
        return;
    }
    const customerName = currentMergeSubmission?.customerName || 'Customer';
    const timestamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
    const fileName = `${customerName.replace(/[^a-zA-Z0-9]/g,'_')}_Merged_${timestamp}.pdf`;
    showLoader('Preparing merged PDF...');
    try { await saveBlobToFolderPicker(mergedPdfBlob, fileName, customerName); }
    finally { hideLoader(); }
}

function viewMergedPDF() {
    if (!mergedPdfBlob || !mergedPdfUrl) {
        showNotification('No merged PDF available','error');
        return;
    }
    const customerName = currentMergeSubmission?.customerName || 'Customer';
    const viewerFileName = document.getElementById('viewerFileName');
    const documentViewer = document.getElementById('documentViewer');
    if (viewerFileName) viewerFileName.textContent = `${customerName} - Merged Documents`;
    if (documentViewer) documentViewer.src = mergedPdfUrl;
    const viewerModal = document.getElementById('viewerModal');
    if (viewerModal) viewerModal.classList.add('active');
    const viewerDownloadBtn = document.getElementById('viewerDownloadBtn');
    if (viewerDownloadBtn) viewerDownloadBtn.style.display='none';
    // Keep merge modal open when viewing merged PDF
    // closeMergeModalFunc();
}

function closeMergeModalFunc() {
    const mergeModal = document.getElementById('mergeModal');
    if (mergeModal) mergeModal.classList.remove('active');
    currentMergeSubmission = null;
    mergeSelectionOrder=[];
    additionalFiles=[];
    additionalFileCounter=0;
    mergedPdfBlob=null;
    if (mergedPdfUrl) { URL.revokeObjectURL(mergedPdfUrl); mergedPdfUrl=null; }
    const mergeDownloadSection = document.getElementById('mergeDownloadSection');
    if (mergeDownloadSection) mergeDownloadSection.style.display='none';
    updateSelectedCount();
}

// ==================== MERGE FUNCTIONS COPY END ====================

function renderRows(submissions) {
    if (!rsaTableBody) return;
    if (submissions.length === 0) {
        rsaTableBody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:40px;color:#666">No records</td></tr>';
        return;
    }
    rsaTableBody.innerHTML = submissions.map(sub=>{
        const docs = (sub.documents||[]).length;
        const viewer = sub.reviewedBy || '-';
        const status = sub.status || '';
        return `<tr>
            <td>${sub.customerName || '-'}</td>
            <td>${viewer}</td>
            <td>${formatTimestamp(sub.uploadedAt)}</td>
            <td>${status.charAt(0).toUpperCase()+status.slice(1)}</td>
            <td>${docs}</td>
        </tr>`;
    }).join('');
}

function loadQueue() {
    // We need to fetch both rsaReady==true records AND rsaSubmitted==true records (old structure)
    // Firestore doesn't support OR in where clauses, so we'll fetch all assigned records and filter in memory
    let q = query(
        collection(db,'submissions'),
        where('assignedToRSA','==', currentUser?.email || ''),
        orderBy('uploadedAt','desc')
    );

    if (typeof unsubscribeQueue === 'function') {
        try { unsubscribeQueue(); } catch (e) { /* ignore */ }
    }

    unsubscribeQueue = onSnapshot(q, async (snap) => {
        const seq = ++queueLoadSeq;
        // Filter for records that are either rsaReady=true OR rsaSubmitted=true (old structure)
        const next = snap.docs
            .map(d => ({ id: d.id, ...d.data() }))
            .filter(s => s.rsaReady === true || s.rsaSubmitted === true);

        const reviewerEmails = [...new Set(next.map(s => s.reviewedBy).filter(Boolean))];
        const uploaderEmails = [...new Set(next.map(s => s.uploadedBy).filter(Boolean))];
        await Promise.all([...reviewerEmails, ...uploaderEmails].map(email => getUserFullName(email)));
        if (seq !== queueLoadSeq) return;

        allSubmissions = next.map(sub => ({
            ...sub,
            reviewedByName: sub.reviewedBy ? (userFullNameCache.get(String(sub.reviewedBy).toLowerCase()) || sub.reviewedBy) : '-',
            uploadedByName: sub.uploadedBy ? (userFullNameCache.get(String(sub.uploadedBy).toLowerCase()) || sub.uploadedBy) : '-'
        }));
        updateApprovedCount();
        renderCurrentTab();
    }, (err) => {
        console.error('loadQueue error', err);
        showNotification('Could not load approved applications: ' + err.message, 'error');
    });
}

function updateApprovedCount() {
    // Exclude both old and new finally submitted records from approved count
    const cnt = allSubmissions.filter(s => (String(s.status || '').toLowerCase() === 'processing_to_pfa' || String(s.status || '').toLowerCase() === 'approved') && !s.finalSubmitted && !s.rsaSubmitted).length;
    if (approvedCountBadge) {
        approvedCountBadge.textContent = cnt;
        approvedCountBadge.style.display = cnt > 0 ? 'inline' : 'none';
    }
    updateFinallySubmittedCount();
}

function switchTab(tabId) {
    console.log('Switching to tab:', tabId);
    currentTab = tabId;
    document.querySelectorAll('.nav-item').forEach(nav=>nav.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach(tab=>tab.classList.remove('active'));
    const targetTab = document.getElementById(`${tabId}Tab`);
    console.log('Target tab element:', targetTab);
    targetTab?.classList.add('active');
    const titles = {
        approved: 'Processing to PFA',
        'finally-submitted': 'Finally Submitted Applications',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'RSA Queue';
    renderCurrentTab();
}

function renderCurrentTab() {
    if (currentTab === 'approved') {
        const start = rsaStartDate.value? new Date(rsaStartDate.value) : null;
        const end = rsaEndDate.value? new Date(rsaEndDate.value) : null;
        if (end) end.setHours(23,59,59,999);
        let list = allSubmissions.filter(s=>{
            if (start && s.uploadedAt && s.uploadedAt.toDate().getTime() < start.getTime()) return false;
            if (end && s.uploadedAt && s.uploadedAt.toDate().getTime() > end.getTime()) return false;
            return true;
        });
        if (rsaSearch && rsaSearch.value.trim()) {
            const qstr = rsaSearch.value.trim().toLowerCase();
            list = list.filter(s => {
                return (s.customerName||'').toLowerCase().includes(qstr) ||
                       (s.reviewedBy||'').toLowerCase().includes(qstr) ||
                       (s.reviewedByName||'').toLowerCase().includes(qstr) ||
                       (s.uploadedBy||'').toLowerCase().includes(qstr) ||
                       (s.uploadedByName||'').toLowerCase().includes(qstr);
            });
        }
        // Exclude both old and new finally submitted records
        list = list.filter(s => {
            const status = String(s.status || '').toLowerCase();
            return (status === 'processing_to_pfa' || status === 'approved') && !s.finalSubmitted && !s.rsaSubmitted;
        });
        renderRsaRows(list);
    } else if (currentTab === 'finally-submitted') {
        renderFinallySubmittedTab();
    }
}

function renderFinallySubmittedTab() {
    const finallySubmittedTableBody = document.getElementById('finallySubmittedTableBody');
    console.log('finallySubmittedTableBody element:', finallySubmittedTableBody);
    if (!finallySubmittedTableBody) {
        console.error('finallySubmittedTableBody not found!');
        return;
    }

    console.log('All submissions:', allSubmissions);
    console.log('Filtering for finally submitted records');
    // Support both old structure (rsaSubmitted=true) and new structure (finalSubmitted=true)
    let list = allSubmissions.filter(s => {
        console.log('Checking submission:', s.customerName, 'finalSubmitted:', s.finalSubmitted, 'rsaSubmitted:', s.rsaSubmitted);
        return s.finalSubmitted === true || s.rsaSubmitted === true;
    });
    console.log('Finally submitted list after filter:', list);

    // Apply filters
    const finalSearch = document.getElementById('finalSearch');
    if (finalSearch && finalSearch.value.trim()) {
        const qstr = finalSearch.value.trim().toLowerCase();
        list = list.filter(s => {
            return (s.customerName||'').toLowerCase().includes(qstr) ||
                   (s.reviewedBy||'').toLowerCase().includes(qstr) ||
                   (s.reviewedByName||'').toLowerCase().includes(qstr) ||
                   (s.uploadedBy||'').toLowerCase().includes(qstr) ||
                   (s.uploadedByName||'').toLowerCase().includes(qstr);
        });
    }

    if (list.length === 0) {
        console.log('No finally submitted records, showing empty message');
        finallySubmittedTableBody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:40px;color:#666">No finally submitted applications</td></tr>';
        return;
    }

    console.log('Rendering', list.length, 'finally submitted records');
    const htmlRows = list.map(sub => {
        const reviewer = sub.reviewedByName || sub.reviewedBy || '-';
        const uploader = sub.uploadedByName || sub.uploadedBy || '-';
        // Support both old (rsaSubmittedBy) and new (finalSubmittedBy) structure
        const rsaOfficer = sub.finalSubmittedBy || sub.rsaSubmittedBy || currentUser?.email || '-';
        const uploadedAt = formatTimestamp(sub.uploadedAt);
        const approvedAt = formatTimestamp(sub.reviewedAt);
        const submittedAt = formatTimestamp(sub.finalSubmittedAt || sub.rsaSubmittedAt || sub.uploadedAt);

        return `<tr>
            <td><strong>${sub.customerName || 'Unknown'}</strong></td>
            <td>${uploader}</td>
            <td>${reviewer}</td>
            <td>${rsaOfficer}</td>
            <td>${uploadedAt}</td>
            <td>${approvedAt}</td>
            <td>${submittedAt}</td>
            <td><span class="status-badge status-approved">Sent to PFA</span></td>
            <td>
                <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')"><i class="fas fa-eye"></i> Details</button>
                <button class="action-btn" onclick="window.downloadAllRsa('${sub.id}')"><i class="fas fa-download"></i> Download All</button>
                <button class="action-btn merge-btn" onclick="window.openMergeModal('${sub.id}')"><i class="fas fa-compress-alt"></i> Merge</button>
                <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button>
            </td>
        </tr>`;
    });

    const joinedHtml = htmlRows.join('');
    console.log('Generated HTML length:', joinedHtml.length);
    console.log('First 200 chars of HTML:', joinedHtml.substring(0, 200));
    finallySubmittedTableBody.innerHTML = joinedHtml;
    console.log('Table body innerHTML set successfully');

    updateFinallySubmittedCount();
}

function updateFinallySubmittedCount() {
    // Support both old structure (rsaSubmitted=true) and new structure (finalSubmitted=true)
    const cnt = allSubmissions.filter(s => s.finalSubmitted === true || s.rsaSubmitted === true).length;
    const badge = document.getElementById('finallySubmittedCount');
    if (badge) {
        badge.textContent = cnt;
        badge.style.display = cnt > 0 ? 'inline' : 'none';
    }
}

function renderRsaRows(submissions) {
    if (!rsaTableBody) return;
    if (submissions.length === 0) {
        rsaTableBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:40px;color:#666">No records</td></tr>';
        return;
    }
    rsaTableBody.innerHTML = submissions.map(sub=>{
        const docs = (sub.documents||[]).length;
        const reviewer = sub.reviewedByName || sub.reviewedBy || '-';
        const uploader = sub.uploadedByName || sub.uploadedBy || '-';
        const approvedAt = formatTimestamp(getApprovedTimestamp(sub));
        const status = String(sub.status || '');
        const statusLabel = (status.toLowerCase() === 'processing_to_pfa' || status.toLowerCase() === 'approved')
            ? 'Processing to PFA'
            : (status.charAt(0).toUpperCase() + status.slice(1));
        const actions = `
            <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')"><i class="fas fa-eye"></i> Details</button>
            <button class="action-btn" onclick="window.downloadAllRsa('${sub.id}')"><i class="fas fa-download"></i> Download All</button>
            <button class="action-btn merge-btn" onclick="window.openMergeModal('${sub.id}')"><i class="fas fa-compress-alt"></i> Merge</button>
            <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button>
            <button class="action-btn" style="background:#10b981;color:#fff;border:none;" onclick="window.finalSubmitRsa('${sub.id}')"><i class="fas fa-paper-plane"></i> Final Submit</button>
        `;
        return `<tr>
            <td>${sub.customerName||'-'}</td>
            <td>${uploader}</td>
            <td>${reviewer}</td>
            <td>${formatTimestamp(sub.uploadedAt)}</td>
            <td>${approvedAt}</td>
            <td>${statusLabel}</td>
            <td>${docs}</td>
            <td>${actions}</td>
        </tr>`;
    }).join('');
}

window.finalSubmitRsa = async (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    const currentStatus = String(sub.status || '').toLowerCase();
    if (!(currentStatus === 'processing_to_pfa' || currentStatus === 'approved')) {
        showNotification('Only applications in Processing to PFA can be finally submitted.', 'warning');
        return;
    }

    const customerName = sub.customerName || 'this customer';
    const confirmed = confirm(`Final submit application for ${customerName}?\n\nThis will remove it from your RSA queue.`);
    if (!confirmed) return;

    try {
        showLoader('Submitting final application...');
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'sent_to_pfa',
            rsaReady: true,
            rsaSubmitted: true,
            rsaSubmittedAt: serverTimestamp(),
            rsaSubmittedBy: currentUser?.email || '',
            finalSubmitted: true,
            finalSubmittedAt: serverTimestamp(),
            finalSubmittedBy: currentUser?.email || ''
        });
        const assignedPayment = await assignPaymentRoundRobin(submissionId);
        if (assignedPayment) {
            try {
                await addDoc(collection(db, 'roundRobinAssignmentsPayment'), {
                    submissionId,
                    customerName: sub.customerName || '',
                    assignedToPayment: assignedPayment,
                    assignedBy: currentUser?.email || '',
                    assignedAt: serverTimestamp()
                });
            } catch (_) {}
        }

        if (sub.uploadedBy) {
            queueUploaderFinalSubmissionEmail({
                submissionId,
                uploaderEmail: sub.uploadedBy,
                customerName: sub.customerName || '',
                rsaEmail: currentUser?.email || ''
            }).catch((emailError) => {
                console.warn('uploader final submission email queue failed:', emailError);
            });
        }
        notifyStatusChangePush({
            currentUser,
            submissionId,
            customerName: sub.customerName || '',
            newStatus: 'sent_to_pfa',
            statusLabel: 'Sent to PFA',
            actionLabel: 'Application Sent to PFA',
            message: `Application for ${sub.customerName || 'this customer'} was finally submitted and sent to PFA.`
        }).catch(() => {});

        // Update the submission in allSubmissions array to reflect the change
        const subIndex = allSubmissions.findIndex(s => s.id === submissionId);
        if (subIndex >= 0) {
            allSubmissions[subIndex] = {
                ...allSubmissions[subIndex],
                status: 'sent_to_pfa',
                rsaSubmitted: true,
                rsaSubmittedAt: new Date(),
                rsaSubmittedBy: currentUser?.email || '',
                finalSubmitted: true,
                finalSubmittedAt: new Date(),
                finalSubmittedBy: currentUser?.email || '',
                assignedToPayment: assignedPayment || ''
            };
        }

        updateApprovedCount();
        updateFinallySubmittedCount();
        renderCurrentTab();
        closeCustomerDetailsModal();

        showNotification(`Final submission recorded. Status set to Sent to PFA${assignedPayment ? ` and assigned to ${assignedPayment}` : ''}.`, 'success');
    } catch (error) {
        console.error('Final submission error:', error);
        showNotification('Final submission failed: ' + (error?.message || 'Unknown error'), 'error');
    } finally {
        hideLoader();
    }
};


rsaFilterBtn?.addEventListener('click', renderCurrentTab);

// search triggers
rsaSearch && rsaSearch.addEventListener('input', renderCurrentTab);

// Finally Submitted tab filters
const finalFilterBtn = document.getElementById('finalFilterBtn');
const finalSearch = document.getElementById('finalSearch');
finalFilterBtn?.addEventListener('click', renderCurrentTab);
finalSearch?.addEventListener('input', renderCurrentTab);

// tab navigation
document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        const tab = item.dataset.tab;
        switchTab(tab);
    });
});

// authentication
auth.onAuthStateChanged(user=>{
    if (!user) return window.location.href = 'index.html';
    currentUser = user;
    loadCurrentRsaProfile(user).then((allowed) => {
        if (!allowed) return;
        loadQueue();
    });
});

window.signOutUser = async () => {
    try { await signOut(auth); } catch (e) { }
    window.location.href = 'index.html';
};

// merge modal controls
document.getElementById('closeMergeModal')?.addEventListener('click', closeMergeModalFunc);
document.getElementById('cancelMerge')?.addEventListener('click', closeMergeModalFunc);
document.getElementById('mergeDocumentsBtn')?.addEventListener('click', handleMergeDocuments);

// Customer details modal controls
document.getElementById('closeCustomerDetails')?.addEventListener('click', closeCustomerDetailsModal);
document.getElementById('closeCustomerDetailsFooter')?.addEventListener('click', closeCustomerDetailsModal);
document.getElementById('closeViewer')?.addEventListener('click', closeViewerModal);
finalSubmitBtn?.addEventListener('click', (e) => {
    e.preventDefault();
    if (currentDetailsSubmissionId) window.finalSubmitRsa(currentDetailsSubmissionId);
});
setupForceRefreshButtons();
setupIdleLogout();

function forceHardRefresh() {
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
}

function setupForceRefreshButtons() {
    document.getElementById('forceRefreshBtn')?.addEventListener('click', (e) => {
        e.preventDefault();
        forceHardRefresh();
    });
    document.getElementById('forceRefreshBtnMobile')?.addEventListener('click', (e) => {
        e.preventDefault();
        forceHardRefresh();
    });
}

function setupIdleLogout() {
    if (idleIntervalHandle) return;
    const bump = () => { idleLastActivity = Date.now(); };
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach((ev) => {
        window.addEventListener(ev, bump, { passive: true });
    });
    window.addEventListener('focus', bump);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) bump(); });

    idleIntervalHandle = setInterval(() => {
        if (!auth.currentUser) return;
        if (Date.now() - idleLastActivity >= IDLE_TIMEOUT_MS) {
            window.signOutUser();
        }
    }, 60 * 1000);
}

// Close modals when clicking outside
window.addEventListener('click', (e) => {
    const customerDetailsModal = document.getElementById('customerDetailsModal');
    const viewerModal = document.getElementById('viewerModal');
    const mergeModal = document.getElementById('mergeModal');
    
    if (e.target === customerDetailsModal) closeCustomerDetailsModal();
    if (e.target === viewerModal) closeViewerModal();
    if (e.target === mergeModal) closeMergeModalFunc();
});

// ==================== CUSTOMER DETAILS MODAL ====================
window.openCustomerDetails = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    currentDetailsSubmissionId = submissionId;
    
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
    setWhatsAppContact(doc.getElementById('detailPhone'), details.phone || '');
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
                        <button class="action-btn" onclick="window.viewDocumentRSA('${submissionId}', ${idx})" title="View">
                            <i class="fas fa-eye"></i>
                        </button>
                        <button class="action-btn" onclick="window.downloadDocumentRSA('${submissionId}', ${idx})" title="Download">
                            <i class="fas fa-download"></i>
                        </button>
                    </div>
                </div>
            `).join('')}
        </div>`;
    
    doc.getElementById('customerDocumentsList').innerHTML = docListHtml;

    if (finalSubmitBtn) {
        const canSubmit = (sub.status || '').toLowerCase() === 'approved';
        finalSubmitBtn.disabled = !canSubmit;
        finalSubmitBtn.style.opacity = canSubmit ? '1' : '0.6';
        finalSubmitBtn.title = canSubmit ? 'Mark as finally submitted' : 'Only approved applications can be submitted';
    }
    
    // Show modal
    const modal = doc.getElementById('customerDetailsModal');
    if (modal) modal.classList.add('active');
};

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (e) {
        return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
}

function closeCustomerDetailsModal() {
    const modal = document.getElementById('customerDetailsModal');
    if (modal) modal.classList.remove('active');
}

window.viewDocumentRSA = (submissionId, docIndex) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub || !sub.documents || !sub.documents[docIndex]) return;
    
    const doc = sub.documents[docIndex];
    const viewerFileName = document.getElementById('viewerFileName');
    const documentViewer = document.getElementById('documentViewer');
    
    if (viewerFileName) viewerFileName.textContent = doc.name;
    if (documentViewer) documentViewer.src = doc.fileUrl;
    
    const viewerModal = document.getElementById('viewerModal');
    if (viewerModal) viewerModal.classList.add('active');
};

window.downloadDocumentRSA = async (submissionId, docIndex) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub || !sub.documents || !sub.documents[docIndex]) return;
    
    const doc = sub.documents[docIndex];
    try {
        showLoader('Downloading document...');
        const response = await fetchWithCorsFallback(doc.fileUrl);
        const blob = await response.blob();
        downloadBlobAsFile(blob, doc.name);
        showNotification('✅ Download started', 'success');
    } catch (error) {
        console.error('Download error:', error);
        showNotification('Download failed: ' + error.message, 'error');
    } finally {
        hideLoader();
    }
};

window.downloadAllRsa = async (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    const docs = sub.documents || [];
    if (!docs.length) {
        showNotification('No documents available for this application', 'warning');
        return;
    }

    const safeCustomerName = (sub.customerName || 'Customer')
        .replace(/[^a-zA-Z0-9\s_-]/g, '_')
        .trim() || 'Customer';

    try {
        showLoader('Preparing document download...');
        let customerFolder = null;

        if ('showDirectoryPicker' in window) {
            showNotification('Select a destination folder to save all documents', 'info');
            const rootFolder = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'downloads' });
            customerFolder = await rootFolder.getDirectoryHandle(safeCustomerName, { create: true });
        }

        for (let i = 0; i < docs.length; i++) {
            const docItem = docs[i];
            if (!docItem?.fileUrl) continue;

            showLoader(`Downloading ${i + 1} of ${docs.length}...`);
            const response = await fetchWithCorsFallback(docItem.fileUrl);
            const blob = await response.blob();
            const fileName = (docItem.name || `${safeCustomerName}_document_${i + 1}.pdf`).replace(/[\\/:*?"<>|]/g, '_');

            if (customerFolder) {
                const fileHandle = await customerFolder.getFileHandle(fileName, { create: true });
                const writable = await fileHandle.createWritable();
                await writable.write(blob);
                await writable.close();
            } else {
                downloadBlobAsFile(blob, fileName);
            }
        }

        showNotification('All documents downloaded successfully', 'success');
    } catch (error) {
        if (error?.name === 'AbortError') {
            showNotification('Download cancelled', 'info');
        } else {
            console.error('Download all error:', error);
            showNotification('Download failed: ' + (error?.message || 'Unknown error'), 'error');
        }
    } finally {
        hideLoader();
    }
};

function closeViewerModal() {
    const modal = document.getElementById('viewerModal');
    if (modal) modal.classList.remove('active');
    const iframe = document.getElementById('documentViewer');
    if (iframe) iframe.src = '';
}

window.signOutUser = () => { window.location.href = 'index.html'; };
window.viewMergedPDF = viewMergedPDF;
