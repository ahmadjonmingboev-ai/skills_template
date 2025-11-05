// WORKSHEET INSTRUCTIONS - JavaScript

// Configuration
const TEMPLATE_PAGE_ID = '3163357481';
const CONFLUENCE_BASE_URL = window.location.origin + '/wiki';

// Trigger Confluence native copy page dialog
function copyTemplatePage() {
    // Construct the copy page URL
    const copyPageUrl = `${CONFLUENCE_BASE_URL}/pages/copy.action?pageId=${TEMPLATE_PAGE_ID}`;
    
    // Open in current window
    window.location.href = copyPageUrl;
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

