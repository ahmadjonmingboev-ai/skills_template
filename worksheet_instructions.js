// WORKSHEET INSTRUCTIONS - JavaScript

// Configuration
const TEMPLATE_PAGE_ID = '3163357481';
const CONFLUENCE_BASE_URL = window.location.origin + '/wiki';

// Copy template page with custom dialogs
async function copyTemplatePage() {
    const button = document.querySelector('.cta-button');
    const originalHTML = button.innerHTML;
    
    try {
        // Disable button
        button.disabled = true;
        button.innerHTML = `
            <div class="spinner"></div>
            <span>Loading...</span>
        `;
        
        // Fetch template page to get default space
        const templateResponse = await fetch(
            `${CONFLUENCE_BASE_URL}/rest/api/content/${TEMPLATE_PAGE_ID}?expand=body.storage,space`,
            { credentials: 'include' }
        );
        
        if (!templateResponse.ok) {
            throw new Error('Failed to fetch template. Please ensure you have view access to the template page.');
        }
        
        const templateData = await templateResponse.json();
        
        // Re-enable button
        button.disabled = false;
        button.innerHTML = originalHTML;
        
        // Show custom input dialog
        const timestamp = new Date().toLocaleString('en-GB', { 
            day: '2-digit', 
            month: '2-digit', 
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        }).replace(/[/:]/g, '-').replace(', ', ' ');
        
        showCopyDialog(templateData, timestamp);
        
    } catch (error) {
        console.error('Init error:', error);
        showToast('Failed to load template: ' + error.message, 'error');
        button.disabled = false;
        button.innerHTML = originalHTML;
    }
}

// Show custom copy dialog
function showCopyDialog(templateData, timestamp) {
    const overlay = document.createElement('div');
    overlay.className = 'input-dialog-overlay';
    overlay.innerHTML = `
        <div class="input-dialog">
            <div class="input-dialog-header">
                <h3>Copy Worksheet Template</h3>
                <p>Configure your new worksheet page</p>
            </div>
            <div class="input-dialog-body">
                <div class="form-group">
                    <label>
                        Page Title <span class="required-indicator">*</span>
                    </label>
                    <input type="text" id="pageTitle" value="My Worksheet - ${timestamp}" required>
                    <div class="form-help">Give your worksheet a meaningful name</div>
                </div>
                
                <div class="form-group">
                    <label>
                        Space Key <span class="required-indicator">*</span>
                    </label>
                    <input type="text" id="spaceKey" value="${templateData.space.key}" placeholder="MYSPACE" required style="text-transform: uppercase;">
                    <div class="form-help">Find this in your space URL: /spaces/<strong>SPACEKEY</strong>/</div>
                </div>
                
                <div class="form-group">
                    <label>
                        Parent Page ID <span class="required-indicator">*</span>
                    </label>
                    <input type="text" id="parentPageId" placeholder="123456789" required>
                    <div class="form-help">The page ID where this worksheet will be created under. Find it in the parent page URL.</div>
                </div>
            </div>
            <div class="input-dialog-footer">
                <button class="dialog-button dialog-button-secondary" onclick="closeCopyDialog()">Cancel</button>
                <button class="dialog-button dialog-button-primary" onclick="executeCopy()">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline; margin-right: 6px;">
                        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                    </svg>
                    Copy Page
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(overlay);
    
    // Store template data for later use
    window.templateDataCache = templateData;
    
    // Focus on first input
    setTimeout(() => {
        document.getElementById('pageTitle').focus();
        document.getElementById('pageTitle').select();
    }, 100);
    
    // Handle Enter key
    overlay.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            executeCopy();
        }
    });
}

// Close copy dialog
function closeCopyDialog() {
    const overlay = document.querySelector('.input-dialog-overlay');
    if (overlay) {
        overlay.remove();
    }
    delete window.templateDataCache;
}

// Execute the copy operation
async function executeCopy() {
    const pageTitle = document.getElementById('pageTitle').value.trim();
    const spaceKey = document.getElementById('spaceKey').value.trim().toUpperCase();
    const parentPageId = document.getElementById('parentPageId').value.trim();
    
    // Validation
    if (!pageTitle) {
        showToast('Please enter a page title', 'error');
        document.getElementById('pageTitle').focus();
        return;
    }
    
    if (!spaceKey) {
        showToast('Please enter a space key', 'error');
        document.getElementById('spaceKey').focus();
        return;
    }
    
    if (!parentPageId) {
        showToast('Please enter a parent page ID', 'error');
        document.getElementById('parentPageId').focus();
        return;
    }
    
    const copyButton = document.querySelector('.input-dialog-footer .dialog-button-primary');
    const originalButtonHTML = copyButton.innerHTML;
    
    try {
        // Disable copy button
        copyButton.disabled = true;
        copyButton.innerHTML = `
            <div class="spinner" style="width: 16px; height: 16px; border-width: 2px;"></div>
            <span style="margin-left: 8px;">Creating page...</span>
        `;
        
        const templateData = window.templateDataCache;
        
        // Create the copy with parent and labels
        const copyPayload = {
            type: 'page',
            title: pageTitle,
            space: {
                key: spaceKey
            },
            ancestors: [
                {
                    id: parentPageId
                }
            ],
            body: {
                storage: {
                    value: templateData.body.storage.value,
                    representation: 'storage'
                }
            },
            metadata: {
                labels: [
                    {
                        prefix: 'global',
                        name: 'hide-title'
                    },
                    {
                        prefix: 'global',
                        name: 'hide-actions'
                    }
                ]
            }
        };
        
        const copyResponse = await fetch(
            `${CONFLUENCE_BASE_URL}/rest/api/content`,
            {
                method: 'POST',
                credentials: 'include',
                headers: {
                    'Content-Type': 'application/json',
                    'X-Atlassian-Token': 'no-check'
                },
                body: JSON.stringify(copyPayload)
            }
        );
        
        if (!copyResponse.ok) {
            const errorData = await copyResponse.json();
            
            // Better error messages
            let errorMessage = errorData.message || 'Failed to create page';
            
            if (errorMessage.includes('permission') || errorMessage.includes('forbidden') || copyResponse.status === 403) {
                errorMessage = `You do not have permission to add pages to space "${spaceKey}". Please ensure the space key is correct or try a different space where you have access.`;
            } else if (errorMessage.includes('not found') || copyResponse.status === 404) {
                errorMessage = `Space "${spaceKey}" or parent page "${parentPageId}" not found. Please check your inputs.`;
            } else if (errorMessage.includes('title') || errorMessage.includes('already exists')) {
                errorMessage = `A page with the title "${pageTitle}" already exists. Please choose a different name.`;
            }
            
            throw new Error(errorMessage);
        }
        
        const newPage = await copyResponse.json();
        
        // Update button
        copyButton.innerHTML = `
            <div class="spinner" style="width: 16px; height: 16px; border-width: 2px;"></div>
            <span style="margin-left: 8px;">Copying attachments...</span>
        `;
        
        // Copy essential attachments (JS and CSS)
        await copyAttachments(TEMPLATE_PAGE_ID, newPage.id);
        
        // Close dialog
        closeCopyDialog();
        
        // Success
        showToast('Page copied successfully! Redirecting to your new worksheet...', 'success');
        
        setTimeout(() => {
            window.location.href = `${CONFLUENCE_BASE_URL}/pages/viewpage.action?pageId=${newPage.id}`;
        }, 1500);
        
    } catch (error) {
        console.error('Copy error:', error);
        showToast(error.message, 'error');
        copyButton.disabled = false;
        copyButton.innerHTML = originalButtonHTML;
    }
}

// Copy essential attachments (only JS and CSS files)
async function copyAttachments(sourcePageId, targetPageId) {
    try {
        // Get template attachments
        const attachmentsResponse = await fetch(
            `${CONFLUENCE_BASE_URL}/rest/api/content/${sourcePageId}/child/attachment`,
            { credentials: 'include' }
        );
        
        if (!attachmentsResponse.ok) {
            console.warn('Could not fetch attachments');
            return;
        }
        
        const attachmentsData = await attachmentsResponse.json();
        
        // Only copy JavaScript and CSS files (not data files)
        const filesToCopy = attachmentsData.results.filter(att => 
            att.title === 'worksheet_optimized.js' || 
            att.title === 'worksheet_optimized.css'
        );
        
        for (const attachment of filesToCopy) {
            try {
                // Download attachment
                const downloadUrl = CONFLUENCE_BASE_URL + attachment._links.download;
                const fileResponse = await fetch(downloadUrl, { credentials: 'include' });
                const fileBlob = await fileResponse.blob();
                
                // Upload to new page
                const formData = new FormData();
                formData.append('file', fileBlob, attachment.title);
                formData.append('comment', 'Copied from template');
                formData.append('minorEdit', 'true');
                
                await fetch(
                    `${CONFLUENCE_BASE_URL}/rest/api/content/${targetPageId}/child/attachment`,
                    {
                        method: 'POST',
                        credentials: 'include',
                        headers: { 'X-Atlassian-Token': 'no-check' },
                        body: formData
                    }
                );
            } catch (err) {
                console.warn(`Failed to copy ${attachment.title}:`, err);
            }
        }
    } catch (error) {
        console.warn('Failed to copy attachments:', error);
    }
}

// Show toast notification
function showToast(message, type = 'success') {
    // Remove existing toasts
    document.querySelectorAll('.toast').forEach(t => t.remove());
    
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    const icon = type === 'success' 
        ? '<svg class="toast-icon" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>'
        : '<svg class="toast-icon" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>';
    
    toast.innerHTML = `
        ${icon}
        <div style="flex: 1; color: var(--text-primary); font-weight: 500;">${message}</div>
    `;
    
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 5000);
}

// Modal functions
function openModal(modalId) {
    document.getElementById(modalId).classList.add('show');
}

function closeModal(modalId) {
    document.getElementById(modalId).classList.remove('show');
}

// Accordion toggle
function toggleAccordion(header) {
    const item = header.parentElement;
    const wasActive = item.classList.contains('active');
    
    // Close all accordions in this modal
    const modal = header.closest('.modal');
    modal.querySelectorAll('.accordion-item').forEach(i => {
        i.classList.remove('active');
    });
    
    // Open this one if it wasn't already active
    if (!wasActive) {
        item.classList.add('active');
    }
}

// Close modal when clicking outside
window.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal')) {
        e.target.classList.remove('show');
    }
});

