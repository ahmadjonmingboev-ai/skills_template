// WORKSHEET INSTRUCTIONS - JavaScript

// Configuration
const TEMPLATE_PAGE_ID = '3163357481';
const CONFLUENCE_BASE_URL = window.location.origin + '/wiki';

// Copy template page programmatically
async function copyTemplatePage() {
    const button = document.querySelector('.cta-button');
    const buttonIcon = button.querySelector('svg');
    const originalHTML = button.innerHTML;
    
    try {
        // Show loading state
        button.disabled = true;
        button.innerHTML = `
            <div class="spinner"></div>
            <span>Copying...</span>
        `;
        
        // Fetch template page content
        const templateResponse = await fetch(
            `${CONFLUENCE_BASE_URL}/rest/api/content/${TEMPLATE_PAGE_ID}?expand=body.storage,space,version,ancestors`,
            { credentials: 'include' }
        );
        
        if (!templateResponse.ok) {
            throw new Error('Failed to fetch template. Please ensure you have view access to the template page.');
        }
        
        const templateData = await templateResponse.json();
        
        // Prompt user for new page title
        const timestamp = new Date().toLocaleString('en-GB', { 
            day: '2-digit', 
            month: '2-digit', 
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        }).replace(/[/:]/g, '-').replace(', ', ' ');
        
        const defaultTitle = `My Worksheet - ${timestamp}`;
        const newTitle = prompt('Enter a name for your new worksheet:', defaultTitle);
        
        if (!newTitle) {
            // User cancelled
            button.disabled = false;
            button.innerHTML = originalHTML;
            return;
        }
        
        // Prompt for space key
        const spaceKey = prompt(
            'Enter your Confluence Space Key (e.g., MYSPACE):\n\nYou can find this in your space URL.',
            templateData.space.key
        );
        
        if (!spaceKey) {
            button.disabled = false;
            button.innerHTML = originalHTML;
            return;
        }
        
        // Create the copy
        const copyPayload = {
            type: 'page',
            title: newTitle,
            space: {
                key: spaceKey.toUpperCase()
            },
            body: {
                storage: {
                    value: templateData.body.storage.value,
                    representation: 'storage'
                }
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
            throw new Error(errorData.message || 'Failed to create page. Please check space key and permissions.');
        }
        
        const newPage = await copyResponse.json();
        
        // Show progress
        button.innerHTML = `
            <div class="spinner"></div>
            <span>Copying attachments...</span>
        `;
        
        // Copy essential attachments (JS and CSS)
        await copyAttachments(TEMPLATE_PAGE_ID, newPage.id);
        
        // Success
        showToast('Page copied successfully! Redirecting to your new worksheet...', 'success');
        
        setTimeout(() => {
            window.location.href = `${CONFLUENCE_BASE_URL}/pages/viewpage.action?pageId=${newPage.id}`;
        }, 1500);
        
    } catch (error) {
        console.error('Copy error:', error);
        showToast('Failed to copy page: ' + error.message, 'error');
        button.disabled = false;
        button.innerHTML = originalHTML;
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

