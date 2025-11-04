// WORKSHEET TOOL - Optimized Version v6 with Performance Fixes
// =============================================================

// ===== SECURITY: HTML ESCAPING (XSS Protection) =====
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;')
        .replace(/\//g, '&#x2F;');
}

// ===== CSV UTILITIES =====
const CSV = {
    parse: function(csvText) {
        const lines = csvText.split('\n').filter(line => line.trim());
        if (lines.length === 0) return [];
        
        const headers = this.parseLine(lines[0]);
        const data = [];
        
        for (let i = 1; i < lines.length; i++) {
            const values = this.parseLine(lines[i]);
            if (values.length === headers.length) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index];
                });
                data.push(row);
            }
        }
        
        return data;
    },
    
    parseLine: function(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                result.push(current);
                current = '';
            } else {
                current += char;
            }
        }
        
        result.push(current);
        return result;
    },
    
    generate: function(data) {
        if (!data || data.length === 0) return '';
        
        const headers = Object.keys(data[0]);
        const lines = [headers.map(h => this.escapeField(h)).join(',')];
        
        data.forEach(row => {
            const values = headers.map(header => this.escapeField(row[header] || ''));
            lines.push(values.join(','));
        });
        
        return lines.join('\n');
    },
    
    escapeField: function(field) {
        const str = String(field);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
    }
};

// ===== TOAST NOTIFICATIONS =====
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    const icons = {
        success: '<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>',
        error: '<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
        warning: '<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>',
        info: '<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>'
    };
    
    toast.innerHTML = `
        ${icons[type]}
        <div class="toast-message">${escapeHtml(message)}</div>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

// ===== CONFIRM MODAL =====
function showConfirm(message, onConfirm) {
    const messageEl = document.getElementById('confirmMessage');
    messageEl.innerHTML = escapeHtml(message);
    document.getElementById('confirmModal').classList.add('show');
    
    const confirmBtn = document.getElementById('confirmButton');
    const newConfirmBtn = confirmBtn.cloneNode(true);
    confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);
    
    newConfirmBtn.onclick = function() {
        closeConfirmModal();
        onConfirm();
    };
}

function closeConfirmModal() {
    document.getElementById('confirmModal').classList.remove('show');
}

// ===== CONFIGURATION =====
const CONFIG = {
    confluenceBaseUrl: '',
    pageId: '',
    worksheetUIFile: 'worksheetui.csv',
    worksheetName: 'Worksheet',
    theme: 'auto',
    authorizedUsers: []
};

// ===== STATE MANAGEMENT =====
const STATE = {
    currentTab: null,
    tabs: [],
    allData: {},
    currentPage: 1,
    pageSize: 25,
    selectedRows: new Set(),
    activeFilters: {},
    tabFilters: {},
    tabSearch: {},
    searchQuery: '',
    currentUser: 'Loading...',
    sortColumn: null,
    sortDirection: null,
    modalData: {},
    rowLocks: {},
    changeHistory: [],
    groupByColumn: null,
    collapsedGroups: new Set(),
    activeUsers: [],
    presenceInterval: null
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', async function() {
    console.log('Worksheet Tool Optimized v6 initializing...');
    await init();
});

async function init() {
    try {
        showLoading(true);
        await getCurrentUser();
        
        // Auto-detect Confluence context before trying to load config
        autoDetectConfluenceContext();
        
        try {
            await loadConfiguration();
            
            // Configuration loaded successfully
            applyTheme(CONFIG.theme);
            document.getElementById('worksheetName').textContent = CONFIG.worksheetName;
            setupEventListeners();
            checkUserAuthorization();
            
            if (STATE.tabs.length > 0) {
                await switchTab(STATE.tabs[0].name);
            }
            
            // Start real-time presence tracking
            startPresenceTracking();
            
            showLoading(false);
        } catch (error) {
            console.warn('Config not found:', error.message);
            // Show welcome message for first-time users
            showLoading(false);
            showFirstTimeSetup();
            return;
        }
    } catch (error) {
        console.error('Init error:', error);
        showToast('Failed to initialize: ' + error.message, 'error');
        showLoading(false);
    }
}

// Clean up on page unload
window.addEventListener('beforeunload', function() {
    stopPresenceTracking();
});

// ===== FIRST-TIME SETUP =====
function showFirstTimeSetup() {
    // Auto-detect Confluence context
    autoDetectConfluenceContext();
    
    // Show welcome overlay
    const overlay = document.createElement('div');
    overlay.className = 'welcome-overlay';
    overlay.innerHTML = `
        <div class="welcome-card">
            <div class="welcome-header">
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                    <polyline points="14 2 14 8 20 8"></polyline>
                    <line x1="12" y1="11" x2="12" y2="17"></line>
                    <polyline points="9 14 12 17 15 14"></polyline>
                </svg>
                <h2>Welcome to Worksheet Tool!</h2>
            </div>
            <div class="welcome-body">
                <p>Get started by configuring your first worksheet.</p>
                <ul>
                    <li>✅ Add worksheet name and settings</li>
                    <li>✅ Create tabs for your data</li>
                    <li>✅ Upload CSV files</li>
                    <li>✅ Configure custom fields</li>
                </ul>
                <p><strong>Note:</strong> Configuration will be saved automatically when you click "Save Configuration".</p>
            </div>
            <div class="welcome-footer">
                <button class="btn btn-primary btn-large" onclick="openConfigModal(); closeWelcomeOverlay();">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <circle cx="12" cy="12" r="3"></circle>
                        <path d="M12 1v6m0 6v6M1 12h6m6 0h6M4.22 4.22l4.24 4.24m5.66 5.66 4.24 4.24m-4.24-14.14 4.24 4.24M4.22 19.78l4.24-4.24"></path>
                    </svg>
                    Start Setup
                </button>
            </div>
        </div>
    `;
    document.body.appendChild(overlay);
    
    // Setup event listeners after DOM is ready
    setupEventListeners();
}

function closeWelcomeOverlay() {
    const overlay = document.querySelector('.welcome-overlay');
    if (overlay) {
        overlay.remove();
    }
}

function autoDetectConfluenceContext() {
    // Only auto-detect Page ID (Base URL is set in CONFIG constant)
    if (!CONFIG.pageId) {
        // Try to get page ID from URL parameters
        const urlParams = new URLSearchParams(window.location.search);
        const pageId = urlParams.get('pageId');
        if (pageId) {
            CONFIG.pageId = pageId;
        } else {
            // Try to extract from pathname (format: /pages/123456789)
            const pathMatch = window.location.pathname.match(/\/pages\/(\d+)/);
            if (pathMatch) {
                CONFIG.pageId = pathMatch[1];
            }
        }
    }
    
    console.log('Confluence context:', {
        baseUrl: CONFIG.confluenceBaseUrl,
        pageId: CONFIG.pageId + (CONFIG.pageId ? ' (auto-detected)' : ' (not found)')
    });
}

// Make function available globally
window.closeWelcomeOverlay = closeWelcomeOverlay;

// ===== USER AUTHORIZATION =====
function checkUserAuthorization() {
    // Authorization logic: Only restrict config button
    // - If authorizedUsers list is empty = everyone can configure
    // - If authorizedUsers list has usernames = only those users can configure
    // Page access is controlled by Confluence page restrictions
    const isAuthorized = CONFIG.authorizedUsers.length === 0 || 
                         CONFIG.authorizedUsers.includes(STATE.currentUsername);
    
    // Only disable config button for unauthorized users
    const configButton = document.getElementById('configButton');
    if (configButton) {
        if (isAuthorized) {
            configButton.disabled = false;
            configButton.style.opacity = '1';
            configButton.title = 'Settings';
        } else {
            configButton.disabled = true;
            configButton.style.opacity = '0.5';
            configButton.style.cursor = 'not-allowed';
            configButton.title = 'Settings (Restricted)';
        }
    }
    
    // Log authorization status for debugging
    console.log(`User Authorization Check:
        Current Username: ${STATE.currentUsername}
        Authorized Users List: ${CONFIG.authorizedUsers.join(', ') || 'Empty (everyone authorized)'}
        Can Configure: ${isAuthorized}`);
}

// ===== CONFLUENCE API =====
async function getCurrentUser() {
    try {
        const response = await fetch(`${CONFIG.confluenceBaseUrl}/rest/api/user/current`, {
            credentials: 'include'
        });
        if (response.ok) {
            const user = await response.json();
            STATE.currentUser = user.displayName || user.username || 'User';
            STATE.currentUsername = user.username || user.key || '';  // Store username separately for authorization
        } else {
            STATE.currentUser = 'Guest';
            STATE.currentUsername = '';
        }
    } catch (error) {
        STATE.currentUser = 'Guest';
        STATE.currentUsername = '';
    }
    document.getElementById('currentUser').textContent = STATE.currentUser;
}

// ===== REAL-TIME PRESENCE =====
async function updatePresence() {
    try {
        const presenceFile = 'worksheet_presence.csv';
        const timestamp = new Date().toISOString();
        
        // Load existing presence data
        let presenceData = [];
        try {
            const presenceText = await fetchAttachment(presenceFile);
            presenceData = CSV.parse(presenceText);
        } catch (error) {
            // File doesn't exist yet
        }
        
        // Remove old entries (older than 10 minutes = 600 seconds)
        // Logic: If someone hasn't updated in 10 minutes, consider them gone
        const tenMinutesAgo = new Date(Date.now() - 600000).toISOString();
        presenceData = presenceData.filter(entry => entry.TIMESTAMP > tenMinutesAgo);
        
        // Update or add current user
        const existingIndex = presenceData.findIndex(entry => entry.USERNAME === STATE.currentUsername);
        if (existingIndex >= 0) {
            presenceData[existingIndex].TIMESTAMP = timestamp;
            presenceData[existingIndex].DISPLAY_NAME = STATE.currentUser;
        } else {
            presenceData.push({
                USERNAME: STATE.currentUsername,
                DISPLAY_NAME: STATE.currentUser,
                TIMESTAMP: timestamp
            });
        }
        
        // Save updated presence
        const csvContent = CSV.generate(presenceData);
        await uploadAttachment(presenceFile, csvContent);
        
        // Update active users list (excluding current user)
        STATE.activeUsers = presenceData
            .filter(entry => entry.USERNAME !== STATE.currentUsername)
            .map(entry => entry.DISPLAY_NAME);
        
        updatePresenceUI();
    } catch (error) {
        console.warn('Failed to update presence:', error);
    }
}

function updatePresenceUI() {
    const userInfoDiv = document.querySelector('.user-info');
    if (!userInfoDiv) return;
    
    // Remove existing viewing icon from user info
    const existingIcon = userInfoDiv.querySelector('.viewing-icon');
    if (existingIcon) {
        existingIcon.remove();
    }
    
    // Add viewing icon next to current user (always show when page is active)
    const viewingIcon = document.createElement('span');
    viewingIcon.className = 'viewing-icon';
    viewingIcon.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
            <circle cx="12" cy="12" r="3"></circle>
        </svg>
    `;
    viewingIcon.title = 'You are viewing this page';
    
    // Insert icon before the username text
    const userNameSpan = userInfoDiv.querySelector('#currentUser');
    if (userNameSpan) {
        userInfoDiv.insertBefore(viewingIcon, userNameSpan);
    }
    
    // Remove existing presence indicator
    let presenceIndicator = document.getElementById('presenceIndicator');
    if (presenceIndicator) {
        presenceIndicator.remove();
    }
    
    // Add new presence indicator if there are other users
    if (STATE.activeUsers.length > 0) {
        presenceIndicator = document.createElement('button');
        presenceIndicator.id = 'presenceIndicator';
        presenceIndicator.className = 'presence-indicator';
        presenceIndicator.innerHTML = `
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <circle cx="12" cy="12" r="3" fill="currentColor"></circle>
            </svg>
            <span>${STATE.activeUsers.length} other${STATE.activeUsers.length > 1 ? 's' : ''} viewing</span>
        `;
        presenceIndicator.title = 'Click to see who\'s viewing';
        presenceIndicator.onclick = openPresenceModal;
        userInfoDiv.parentNode.insertBefore(presenceIndicator, userInfoDiv);
    }
}

function openPresenceModal() {
    const modal = document.getElementById('presenceModal');
    if (!modal) {
        // Create modal if it doesn't exist
        createPresenceModal();
    }
    
    const container = document.getElementById('presenceModalContent');
    container.innerHTML = '';
    
    if (STATE.activeUsers.length === 0) {
        container.innerHTML = '<p style="text-align: center; padding: 20px; color: var(--text-muted);">No other viewers</p>';
    } else {
        const userList = document.createElement('div');
        userList.className = 'viewer-list';
        
        STATE.activeUsers.forEach(userName => {
            const userItem = document.createElement('div');
            userItem.className = 'viewer-item';
            userItem.innerHTML = `
                <div class="viewer-avatar">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                        <circle cx="12" cy="7" r="4"></circle>
                    </svg>
                </div>
                <div class="viewer-info">
                    <div class="viewer-name">${escapeHtml(userName)}</div>
                    <div class="viewer-status">
                        <span class="status-dot"></span>
                        Viewing now
                    </div>
                </div>
            `;
            userList.appendChild(userItem);
        });
        
        container.appendChild(userList);
    }
    
    document.getElementById('presenceModal').classList.add('show');
}

function closePresenceModal() {
    document.getElementById('presenceModal').classList.remove('show');
}

function createPresenceModal() {
    const modal = document.createElement('div');
    modal.id = 'presenceModal';
    modal.className = 'modal';
    modal.innerHTML = `
        <div class="modal-content modal-small">
            <div class="modal-header">
                <h2>
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                        <circle cx="9" cy="7" r="4"></circle>
                        <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                        <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                    </svg>
                    Active Viewers
                </h2>
                <button class="modal-close" onclick="closePresenceModal()">×</button>
            </div>
            <div class="modal-body" id="presenceModalContent"></div>
        </div>
    `;
    document.body.appendChild(modal);
}

function startPresenceTracking() {
    // Update immediately
    updatePresence();
    
    // Update every 5 minutes (300000 ms)
    STATE.presenceInterval = setInterval(updatePresence, 300000);
}

function stopPresenceTracking() {
    if (STATE.presenceInterval) {
        clearInterval(STATE.presenceInterval);
        STATE.presenceInterval = null;
    }
}

async function fetchAttachment(filename) {
    const url = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    const response = await fetch(url, { credentials: 'include' });
    
    if (!response.ok) {
        throw new Error(`Failed to fetch attachments: ${response.status}`);
    }
    
    const data = await response.json();
    const attachment = data.results.find(att => att.title === filename);
    
    if (!attachment) {
        throw new Error(`Attachment ${filename} not found`);
    }
    
    const downloadUrl = CONFIG.confluenceBaseUrl + attachment._links.download;
    const fileResponse = await fetch(downloadUrl, { credentials: 'include' });
    
    if (!fileResponse.ok) {
        throw new Error(`Failed to download ${filename}`);
    }
    
    return await fileResponse.text();
}

async function uploadAttachment(filename, content) {
    const attachmentsUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    const attachmentsResponse = await fetch(attachmentsUrl, { credentials: 'include' });
    
    if (!attachmentsResponse.ok) {
        throw new Error(`Failed to fetch attachments`);
    }
    
    const attachmentsData = await attachmentsResponse.json();
    const existingAttachment = attachmentsData.results.find(att => att.title === filename);
    
    const formData = new FormData();
    const blob = new Blob([content], { type: 'text/csv' });
    formData.append('file', blob, filename);
    formData.append('comment', `Updated by ${STATE.currentUser}`);
    formData.append('minorEdit', 'true');
    
    let uploadUrl;
    if (existingAttachment) {
        uploadUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment/${existingAttachment.id}/data`;
    } else {
        uploadUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    }
    
    const response = await fetch(uploadUrl, {
        method: 'POST',
        credentials: 'include',
        headers: { 'X-Atlassian-Token': 'no-check' },
        body: formData
    });
    
    if (!response.ok) {
        throw new Error(`Upload failed: ${response.status}`);
    }
    
    return await response.json();
}

// ===== FILE UPLOAD HANDLER - INDEPENDENT UPLOAD =====
async function handleFileUpload(tabIndex, file) {
    if (!file) return;
    
    try {
        showLoading(true);
        
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const content = e.target.result;
                
                // Upload file with its original name directly to attachments
                // No dependency on tab configuration
                const filename = file.name;
                
                await uploadAttachment(filename, content);
                
                showToast(`File "${filename}" uploaded successfully to page attachments`, 'success');
                
                // If a tab exists and matches, suggest setting it
                const tab = STATE.tabs[tabIndex];
                if (tab && !tab.filename) {
                    // Auto-populate the filename field for convenience
                    const filenameInput = document.querySelector(`#tabsConfigContainer .tab-config-item:nth-child(${tabIndex + 1}) .tab-filename-input`);
                    if (filenameInput) {
                        filenameInput.value = filename;
                        showToast(`Tip: Filename set to "${filename}". Click Save Configuration to apply.`, 'info');
                    }
                }
            } catch (error) {
                console.error('Upload error:', error);
                showToast('Failed to upload file: ' + error.message, 'error');
            } finally {
                showLoading(false);
            }
        };
        
        reader.onerror = function() {
            showToast('Failed to read file', 'error');
            showLoading(false);
        };
        
        reader.readAsText(file);
    } catch (error) {
        console.error('File upload error:', error);
        showToast('Failed to upload file: ' + error.message, 'error');
        showLoading(false);
    }
}

// ===== CONFIGURATION - FIXED GROUP_BY SAVING =====
async function loadConfiguration() {
    const configText = await fetchAttachment(CONFIG.worksheetUIFile);
    const configData = CSV.parse(configText);
    
    STATE.tabs = [];
    
    configData.forEach(row => {
        const field = row.Fields || row.fields;
        const value = row.Values || row.values;
        
        if (!field || value === undefined || value === '') return;
        
        if (field === 'WORKSHEET_NAME') {
            CONFIG.worksheetName = value;
        } else if (field === 'CONFLUENCE_BASE_URL') {
            CONFIG.confluenceBaseUrl = value;
        } else if (field === 'PAGEID') {
            CONFIG.pageId = value;
        } else if (field === 'THEME') {
            CONFIG.theme = value;
        } else if (field === 'AUTHORIZED_USERS') {
            CONFIG.authorizedUsers = value.split(',').map(u => u.trim()).filter(u => u);
        } else if (field.match(/^TAB_FILE_\d+$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            
            if (!STATE.tabs.find(t => t.key === tabKey)) {
                STATE.tabs.push({
                    key: tabKey,
                    name: value.replace('.csv', ''),
                    displayName: value,
                    filename: '',
                    primaryKey: 'ID',
                    sumColumn: '',
                    writeBack: false,
                    customColumns: [],
                    columnOrder: [],
                    groupByOptions: []
                });
            }
        } else if (field.match(/^TAB_FILE_\d+_NAME$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.filename = value;
        } else if (field.match(/^TAB_FILE_\d+_PRIMARY_KEY$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.primaryKey = value;
        } else if (field.match(/^TAB_FILE_\d+_SUM_COLUMN$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.sumColumn = value;
        } else if (field.match(/^TAB_FILE_\d+_WRITE_BACK$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.writeBack = value.toUpperCase() === 'TRUE';
        } else if (field.match(/^TAB_FILE_\d+_CUSTOM_COLUMNS$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab && value) {
                try {
                    tab.customColumns = JSON.parse(value);
                } catch (e) {
                    console.warn('Failed to parse custom columns:', e);
                }
            }
        } else if (field.match(/^TAB_FILE_\d+_COLUMN_ORDER$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab && value) {
                tab.columnOrder = value.split(',').map(c => c.trim());
            }
        } else if (field.match(/^TAB_FILE_\d+_GROUP_BY$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab && value) {
                tab.groupByOptions = value.split(',').map(c => c.trim());
            }
        }
    });
    
    renderTabs();
}

async function saveConfiguration() {
    try {
        showLoading(true);
        
        CONFIG.worksheetName = document.getElementById('configWorksheetName').value.trim();
        CONFIG.confluenceBaseUrl = document.getElementById('configBaseUrl').value.trim();
        CONFIG.pageId = document.getElementById('configPageId').value.trim();
        CONFIG.theme = document.getElementById('configTheme').value;
        CONFIG.authorizedUsers = document.getElementById('configAuthorizedUsers').value
            .split(',').map(u => u.trim()).filter(u => u);
        
        const tabConfigs = document.querySelectorAll('.tab-config-item');
        STATE.tabs = [];
        
        // Validation for required fields
        let validationErrors = [];
        
        for (let index = 0; index < tabConfigs.length; index++) {
            const item = tabConfigs[index];
            const displayName = item.querySelector('.tab-display-input').value.trim();
            const filename = item.querySelector('.tab-filename-input').value.trim();
            const primaryKey = item.querySelector('.tab-primary-key-input').value.trim() || 'ID';
            const sumColumn = item.querySelector('.tab-sum-input').value.trim();
            const writeBack = item.querySelector('.tab-writeback-input').checked;
            const columnOrder = item.querySelector('.tab-columns-input').value.trim();
            const groupByOptions = item.querySelector('.tab-groupby-input').value.trim();
            
            // Validate required fields
            if (!displayName) {
                validationErrors.push(`Tab ${index + 1}: Display Name is required`);
            }
            if (!filename) {
                validationErrors.push(`Tab ${index + 1}: CSV Filename is required`);
            }
            
            // Get custom columns
            const customColumns = [];
            item.querySelectorAll('.custom-column-config').forEach(colConfig => {
                const colName = colConfig.querySelector('.custom-col-name').value.trim();
                const colType = colConfig.querySelector('.custom-col-type').value;
                const colRequired = colConfig.querySelector('.custom-col-required').checked;
                const colOptions = colConfig.querySelector('.custom-col-options').value.trim();
                const colColor = colConfig.querySelector('.custom-col-color')?.value || '';
                
                if (colName) {
                    const customCol = {
                        name: colName,
                        type: colType,
                        required: colRequired,
                        options: colOptions ? colOptions.split(',').map(o => o.trim()) : []
                    };
                    
                    // Store color mapping for dropdown options
                    if (colType === 'dropdown' && colOptions) {
                        const optionsList = colOptions.split(',').map(o => o.trim());
                        customCol.colorMapping = {};
                        
                        // Get color selections for each option
                        const colorSelects = colConfig.querySelectorAll('.option-color-select');
                        colorSelects.forEach((select, index) => {
                            if (optionsList[index]) {
                                customCol.colorMapping[optionsList[index]] = select.value || 'default';
                            }
                        });
                    }
                    
                    customColumns.push(customCol);
                }
            });
            
            if (displayName && filename) {
                const tab = {
                    key: `TAB_${index + 1}`,
                    name: filename.replace('.csv', ''),
                    displayName: displayName,
                    filename: filename,
                    primaryKey: primaryKey,
                    sumColumn: sumColumn,
                    writeBack: writeBack,
                    customColumns: customColumns,
                    columnOrder: columnOrder ? columnOrder.split(',').map(c => c.trim()) : [],
                    groupByOptions: groupByOptions ? groupByOptions.split(',').map(c => c.trim()) : []
                };
                
                STATE.tabs.push(tab);
                
                // Auto-create custom columns file if needed
                if (customColumns.length > 0) {
                    await createCustomColumnsFile(tab);
                }
            }
        }
        
        // Check if there are validation errors
        if (validationErrors.length > 0) {
            showToast(validationErrors.join('\n'), 'error');
            showLoading(false);
            return;
        }
        
        // Check if at least one tab is configured
        if (STATE.tabs.length === 0) {
            showToast('At least one tab must be configured with Display Name and CSV Filename', 'error');
            showLoading(false);
            return;
        }
        
        const configRows = [
            { Fields: 'WORKSHEET_NAME', Values: CONFIG.worksheetName },
            { Fields: 'CONFLUENCE_BASE_URL', Values: CONFIG.confluenceBaseUrl },
            { Fields: 'PAGEID', Values: CONFIG.pageId },
            { Fields: 'THEME', Values: CONFIG.theme },
            { Fields: 'AUTHORIZED_USERS', Values: CONFIG.authorizedUsers.join(',') }
        ];
        
        STATE.tabs.forEach((tab, index) => {
            configRows.push({ Fields: `TAB_FILE_${index + 1}`, Values: tab.displayName });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_NAME`, Values: tab.filename });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_PRIMARY_KEY`, Values: tab.primaryKey });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_SUM_COLUMN`, Values: tab.sumColumn });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_WRITE_BACK`, Values: tab.writeBack ? 'TRUE' : 'FALSE' });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_CUSTOM_COLUMNS`, Values: JSON.stringify(tab.customColumns) });
            if (tab.columnOrder.length > 0) {
                configRows.push({ Fields: `TAB_FILE_${index + 1}_COLUMN_ORDER`, Values: tab.columnOrder.join(',') });
            }
            if (tab.groupByOptions.length > 0) {
                configRows.push({ Fields: `TAB_FILE_${index + 1}_GROUP_BY`, Values: tab.groupByOptions.join(',') });
            }
        });
        
        const csvContent = CSV.generate(configRows);
        await uploadAttachment(CONFIG.worksheetUIFile, csvContent);
        
        // Check if this was first-time setup
        const isFirstTime = document.querySelector('.welcome-overlay') !== null;
        
        if (isFirstTime) {
            showToast('Configuration saved! Refreshing page...', 'success');
            closeConfigModal();
            // Refresh page to load with new config
            setTimeout(() => {
                window.location.reload();
            }, 1500);
        } else {
            showToast('Configuration saved!', 'success');
            closeConfigModal();
            await init();
        }
        
    } catch (error) {
        console.error('Save config error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== CUSTOM COLUMNS FILE - FIXED PRIMARY KEY =====
async function createCustomColumnsFile(tab) {
    const customFileName = `${tab.name}_custom_columns.csv`;
    
    let existingData = [];
    try {
        const existingText = await fetchAttachment(customFileName);
        existingData = CSV.parse(existingText);
    } catch (error) {
        console.log('Creating new custom columns file:', customFileName);
    }
    
    // Use tab's primary key, not hardcoded
    if (existingData.length === 0) {
        const headers = [tab.primaryKey, 'UPDATED_BY', 'UPDATED_TIME'];
        tab.customColumns.forEach(col => {
            headers.push(col.name);
        });
        
        const emptyRow = {};
        headers.forEach(h => emptyRow[h] = '');
        existingData = [emptyRow];
    }
    
    const csvContent = CSV.generate(existingData);
    await uploadAttachment(customFileName, csvContent);
}

// ===== TAB MANAGEMENT =====
function renderTabs() {
    const container = document.getElementById('tabNavigation');
    container.innerHTML = '';
    
    STATE.tabs.forEach(tab => {
        const button = document.createElement('button');
        button.className = 'tab-button';
        button.dataset.tabName = tab.name;
        button.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
            </svg>
            <span>${escapeHtml(tab.displayName)}</span>
        `;
        button.onclick = () => switchTab(tab.name);
        container.appendChild(button);
    });
}

async function switchTab(tabName) {
    try {
        showLoading(true);
        
        STATE.currentTab = tabName;
        STATE.currentPage = 1;
        STATE.selectedRows.clear();
        STATE.sortColumn = null;
        STATE.sortDirection = null;
        STATE.groupByColumn = null;
        STATE.collapsedGroups.clear();
        
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tabName === tabName);
        });
        
        if (!STATE.allData[tabName]) {
            await loadTabData(tabName);
        }
        
        const filters = STATE.tabFilters[tabName] || {};
        const search = STATE.tabSearch[tabName] || '';
        
        STATE.activeFilters = filters;
        STATE.searchQuery = search;
        document.getElementById('searchInput').value = search;
        
        const activeFilterCount = Object.keys(filters).length;
        const badge = document.getElementById('filterBadge');
        if (activeFilterCount > 0) {
            badge.textContent = activeFilterCount;
            badge.style.display = 'inline-flex';
        } else {
            badge.style.display = 'none';
        }
        
        applyFilters();
        renderTable();
        updateSummary();
        updateButtonVisibility();
        updateGroupByOptions();
        
        showLoading(false);
    } catch (error) {
        console.error('Switch tab error:', error);
        showToast('Failed to load tab: ' + error.message, 'error');
        showLoading(false);
    }
}

async function loadTabData(tabName) {
    const tab = STATE.tabs.find(t => t.name === tabName);
    if (!tab) return;
    
    const dataText = await fetchAttachment(tab.filename);
    const data = CSV.parse(dataText);
    
    let customColumnsData = [];
    const customFileName = `${tab.name}_custom_columns.csv`;
    
    if (tab.customColumns && tab.customColumns.length > 0) {
        try {
            const customText = await fetchAttachment(customFileName);
            customColumnsData = CSV.parse(customText);
        } catch (error) {
            console.warn('Custom columns file not found:', customFileName);
            await createCustomColumnsFile(tab);
        }
    }
    
    STATE.allData[tabName] = {
        rawData: data,
        filteredData: [...data],
        customColumnsData: customColumnsData,
        customColumns: tab.customColumns || [],
        primaryKey: tab.primaryKey
    };
}

// ===== SORTING - Optimized =====
function sortColumn(column) {
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    if (!tabData) return;
    
    if (STATE.sortColumn === column) {
        if (STATE.sortDirection === 'asc') {
            STATE.sortDirection = 'desc';
        } else if (STATE.sortDirection === 'desc') {
            STATE.sortColumn = null;
            STATE.sortDirection = null;
        } else {
            STATE.sortDirection = 'asc';
        }
    } else {
        STATE.sortColumn = column;
        STATE.sortDirection = 'asc';
    }
    
    if (STATE.sortColumn && STATE.sortDirection) {
        tabData.filteredData.sort((a, b) => {
            let aVal, bVal;
            
            const isCustomColumn = tab.customColumns.some(c => c.name === column);
            
            if (isCustomColumn) {
                const aCustom = tabData.customColumnsData.find(c => c[tab.primaryKey] == a[tab.primaryKey]);
                const bCustom = tabData.customColumnsData.find(c => c[tab.primaryKey] == b[tab.primaryKey]);
                aVal = aCustom ? aCustom[column] : '';
                bVal = bCustom ? bCustom[column] : '';
            } else {
                aVal = a[column] || '';
                bVal = b[column] || '';
            }
            
            const aNum = parseFloat(aVal);
            const bNum = parseFloat(bVal);
            
            if (!isNaN(aNum) && !isNaN(bNum)) {
                return STATE.sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
            }
            
            aVal = String(aVal).toLowerCase();
            bVal = String(bVal).toLowerCase();
            
            if (STATE.sortDirection === 'asc') {
                return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
            } else {
                return aVal > bVal ? -1 : aVal < bVal ? 1 : 0;
            }
        });
    }
    
    STATE.currentPage = 1;
    renderTable();
}

// ===== GROUPING =====
function groupBy(column) {
    STATE.groupByColumn = column;
    STATE.collapsedGroups.clear();
    STATE.currentPage = 1;
    renderTable();
}

function toggleGroup(groupName) {
    if (STATE.collapsedGroups.has(groupName)) {
        STATE.collapsedGroups.delete(groupName);
    } else {
        STATE.collapsedGroups.add(groupName);
    }
    renderTable();
}

function updateGroupByOptions() {
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const container = document.getElementById('groupBySelect');
    if (!container) return;
    
    container.innerHTML = '<option value="">No Grouping</option>';
    
    if (tab && tab.groupByOptions) {
        tab.groupByOptions.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            container.appendChild(option);
        });
    }
    
    container.value = STATE.groupByColumn || '';
}

// ===== TABLE RENDERING - Optimized without action column =====
function renderTable() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const data = tabData.filteredData;
    const customColumnsData = tabData.customColumnsData || [];
    const customColumns = tab.customColumns || [];
    
    const allDataColumns = data.length > 0 ? Object.keys(data[0]) : [];
    
    let orderedDataColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        orderedDataColumns = tab.columnOrder.filter(col => allDataColumns.includes(col));
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedDataColumns = [...orderedDataColumns, ...remainingColumns];
    } else {
        orderedDataColumns = allDataColumns;
    }
    
    const customColumnNames = customColumns.map(c => c.name);
    
    // Add system columns (UPDATED_BY, UPDATED_TIME) at the END if custom columns exist
    const systemColumns = [];
    if (customColumns.length > 0) {
        systemColumns.push('UPDATED_BY', 'UPDATED_TIME');
    }
    
    const allCustomColumns = [...customColumnNames, ...systemColumns];
    const allColumns = [...customColumnNames, ...orderedDataColumns, ...systemColumns];
    
    let displayData = data;
    let groups = {};
    
    if (STATE.groupByColumn) {
        data.forEach(row => {
            let groupValue;
            
            // Check if it's a custom column
            if (customColumnNames.includes(STATE.groupByColumn)) {
                const customRow = customColumnsData.find(c => c[tab.primaryKey] == row[tab.primaryKey]);
                groupValue = customRow ? (customRow[STATE.groupByColumn] || 'Ungrouped') : 'Ungrouped';
            } else {
                groupValue = row[STATE.groupByColumn] || 'Ungrouped';
            }
            
            if (!groups[groupValue]) {
                groups[groupValue] = [];
            }
            groups[groupValue].push(row);
        });
    }
    
    const thead = document.getElementById('tableHeader');
    thead.innerHTML = `
        <tr>
            <th class="checkbox-col"><input type="checkbox" id="headerCheckbox" class="row-checkbox"></th>
            ${allColumns.map((col) => {
                const sortClass = STATE.sortColumn === col 
                    ? (STATE.sortDirection === 'asc' ? 'sort-asc' : 'sort-desc')
                    : 'sortable';
                return `<th class="${sortClass}" data-column="${escapeHtml(col)}">
                    ${escapeHtml(col)}
                    <div class="resize-handle" data-column="${escapeHtml(col)}"></div>
                </th>`;
            }).join('')}
        </tr>
    `;
    
    thead.querySelectorAll('.sortable, .sort-asc, .sort-desc').forEach(th => {
        th.addEventListener('click', function(e) {
            // Don't sort if clicking on resize handle
            if (e.target.classList.contains('resize-handle')) return;
            
            const column = this.dataset.column;
            if (column) {
                sortColumn(column);
            }
        });
    });
    
    // Add resize functionality
    setupColumnResize();
    
    document.getElementById('headerCheckbox').addEventListener('change', function() {
        if (this.checked) {
            selectAllRows();
        } else {
            unselectAllRows();
        }
    });
    
    const totalRows = data.length;
    const pageSize = parseInt(STATE.pageSize);
    const totalPages = Math.ceil(totalRows / pageSize);
    const startIndex = (STATE.currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalRows);
    const pageData = data.slice(startIndex, endIndex);
    
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (pageData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="${allColumns.length + 1}" style="text-align: center; padding: 40px;">
                    No data available
                </td>
            </tr>
        `;
        renderPagination(totalPages, startIndex, endIndex, totalRows);
        return;
    }
    
    if (STATE.groupByColumn && Object.keys(groups).length > 0) {
        Object.entries(groups).forEach(([groupName, groupRows]) => {
            const isCollapsed = STATE.collapsedGroups.has(groupName);
            
            const groupHeader = document.createElement('tr');
            groupHeader.className = 'group-header';
            groupHeader.innerHTML = `
                <td colspan="${allColumns.length + 1}">
                    <button class="group-toggle-btn" onclick="toggleGroup('${escapeHtml(groupName)}')">
                        ${isCollapsed ? '▶' : '▼'} ${escapeHtml(groupName)} (${groupRows.length} items)
                    </button>
                </td>
            `;
            tbody.appendChild(groupHeader);
            
            if (!isCollapsed) {
                groupRows.slice(startIndex, Math.min(startIndex + pageSize, groupRows.length)).forEach((row, index) => {
                    renderRow(row, startIndex + index, tbody, allColumns, customColumnsData, tab, customColumnNames, allCustomColumns);
                });
            }
        });
    } else {
        pageData.forEach((row, index) => {
            renderRow(row, startIndex + index, tbody, allColumns, customColumnsData, tab, customColumnNames, allCustomColumns);
        });
    }
    
    renderPagination(totalPages, startIndex, endIndex, totalRows);
}

function renderRow(row, globalIndex, tbody, allColumns, customColumnsData, tab, customColumnNames, allCustomColumns) {
    const primaryKeyValue = row[tab.primaryKey];
    const isSelected = STATE.selectedRows.has(globalIndex);
    const isLocked = STATE.rowLocks[primaryKeyValue];
    
    const tr = document.createElement('tr');
    tr.className = isSelected ? 'selected' : '';
    if (isLocked) tr.className += ' locked';
    tr.dataset.index = globalIndex;
    
    const checkboxTd = document.createElement('td');
    checkboxTd.className = 'checkbox-col';
    checkboxTd.innerHTML = `<input type="checkbox" class="row-checkbox" ${isSelected ? 'checked' : ''} ${isLocked ? 'disabled' : ''}>`;
    checkboxTd.querySelector('.row-checkbox').addEventListener('change', function() {
        toggleRowSelection(globalIndex, this.checked);
    });
    tr.appendChild(checkboxTd);
    
    // Get custom row data once
    const customRow = customColumnsData.find(c => c[tab.primaryKey] == primaryKeyValue);
    
    // Render custom columns first (excluding system columns)
    customColumnNames.forEach(colName => {
        const td = document.createElement('td');
        const customCol = tab.customColumns.find(c => c.name === colName);
        const value = customRow ? (customRow[colName] || '') : '';
        
        // Apply status colors for dropdowns
        if (customCol && customCol.type === 'dropdown' && value) {
            let statusClass = '';
            
            // Check if custom color mapping exists
            if (customCol.colorMapping && customCol.colorMapping[value]) {
                const colorName = customCol.colorMapping[value];
                if (colorName !== 'default') {
                    statusClass = `status-color-${colorName}`;
                } else {
                    statusClass = getStatusClass(value);
                }
            } else {
                statusClass = getStatusClass(value);
            }
            
            td.innerHTML = `<span class="status-badge ${statusClass}">${escapeHtml(value)}</span>`;
        } else {
            td.textContent = value;
        }
        
        tr.appendChild(td);
    });
    
    // Render data columns
    allColumns.forEach(col => {
        if (!allCustomColumns.includes(col)) {
            const td = document.createElement('td');
            const value = row[col];
            
            if (col === tab.sumColumn && !isNaN(value)) {
                td.textContent = formatNumber(value);
                td.style.textAlign = 'right';
            } else {
                td.textContent = value !== undefined && value !== null ? value : '';
            }
            
            tr.appendChild(td);
        }
    });
    
    // Render system columns at the END (UPDATED_BY, UPDATED_TIME)
    if (tab.customColumns && tab.customColumns.length > 0) {
        ['UPDATED_BY', 'UPDATED_TIME'].forEach(colName => {
            const td = document.createElement('td');
            const value = customRow ? (customRow[colName] || '') : '';
            td.textContent = value;
            td.style.fontStyle = 'italic';
            td.style.color = 'var(--text-muted)';
            td.style.fontSize = '13px';
            tr.appendChild(td);
        });
    }
    
    tbody.appendChild(tr);
}

function getStatusClass(value) {
    const val = value.toLowerCase();
    if (val.includes('open')) return 'status-open';
    if (val.includes('closed')) return 'status-closed';
    if (val.includes('pending')) return 'status-pending';
    if (val.includes('active')) return 'status-active';
    if (val.includes('inactive')) return 'status-inactive';
    return '';
}

// ===== DUPLICATE ROW =====
function duplicateRow(index) {
    if (STATE.selectedRows.size !== 1) {
        showToast('Please select exactly one row to duplicate', 'warning');
        return;
    }
    
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const selectedIndex = Array.from(STATE.selectedRows)[0];
    const row = tabData.filteredData[selectedIndex];
    
    if (!row) return;
    
    const newRow = { ...row };
    newRow[tab.primaryKey] = `${row[tab.primaryKey]}_COPY_${Date.now()}`;
    
    openEditRowModal(newRow, true);
}

// ===== ROW LOCKING =====
async function lockRow(primaryKey) {
    STATE.rowLocks[primaryKey] = {
        user: STATE.currentUser,
        timestamp: new Date().toISOString()
    };
    return true;
}

async function unlockRow(primaryKey) {
    delete STATE.rowLocks[primaryKey];
    return true;
}

// ===== CHANGE HISTORY =====
function trackChange(action, tabName, primaryKey, oldValue, newValue) {
    const change = {
        id: Date.now(),
        timestamp: new Date().toISOString(),
        user: STATE.currentUser,
        action: action,
        tab: tabName,
        primaryKey: primaryKey,  // This should be the actual ID value, not the column name
        oldValue: oldValue,
        newValue: newValue
    };
    
    STATE.changeHistory.push(change);
    
    if (STATE.changeHistory.length > 1000) {
        STATE.changeHistory = STATE.changeHistory.slice(-1000);
    }
    
    saveChangeHistory();
}

async function saveChangeHistory() {
    try {
        const historyFile = `${STATE.currentTab}_change_history.csv`;
        const csvContent = CSV.generate(STATE.changeHistory);
        await uploadAttachment(historyFile, csvContent);
    } catch (error) {
        console.warn('Failed to save change history:', error);
    }
}

// ===== MODAL PERSISTENCE =====
function saveModalState(modalId, data) {
    STATE.modalData[modalId] = data;
}

function loadModalState(modalId) {
    return STATE.modalData[modalId] || {};
}

function clearModalState(modalId) {
    delete STATE.modalData[modalId];
}

// ===== COMMENTS/CUSTOM FIELDS MODAL =====
function openCommentsModal() {
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    
    if (!tabData) return;
    
    const selectedCount = STATE.selectedRows.size;
    if (selectedCount === 0) {
        showToast('Please select at least one row', 'warning');
        return;
    }
    
    if (!tab.customColumns || tab.customColumns.length === 0) {
        showToast('No custom columns configured for this tab', 'error');
        return;
    }
    
    document.getElementById('commentRowCount').textContent = selectedCount;
    
    const container = document.getElementById('commentFormContainer');
    container.innerHTML = '';
    
    const modalState = loadModalState('commentModal');
    
    const selectedIndices = Array.from(STATE.selectedRows);
    const selectedRows = selectedIndices.map(idx => tabData.filteredData[idx]);
    const primaryKeys = selectedRows.map(row => row[tab.primaryKey]);
    
    tab.customColumns.forEach(col => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const existingValues = [];
        primaryKeys.forEach(pk => {
            const customRow = tabData.customColumnsData.find(c => c[tab.primaryKey] == pk);
            if (customRow && customRow[col.name]) {
                existingValues.push(customRow[col.name]);
            }
        });
        
        const uniqueExisting = [...new Set(existingValues)];
        const previousValue = uniqueExisting.length === 1 ? uniqueExisting[0] : 
                             uniqueExisting.length > 1 ? '(Multiple values)' : '';
        
        const savedValue = modalState[col.name] || previousValue || '';
        
        let inputHtml = '';
        
        if (col.type === 'dropdown' && col.options.length > 0) {
            inputHtml = `
                <select class="form-control comment-field" data-column="${escapeHtml(col.name)}" 
                        ${col.required ? 'required' : ''}>
                    <option value="">Select...</option>
                    ${col.options.map(opt => `
                        <option value="${escapeHtml(opt)}" ${savedValue === opt ? 'selected' : ''}>
                            ${escapeHtml(opt)}
                        </option>
                    `).join('')}
                </select>
            `;
        } else if (col.type === 'date') {
            inputHtml = `
                <input type="date" class="form-control comment-field" 
                       data-column="${escapeHtml(col.name)}" 
                       value="${escapeHtml(savedValue)}"
                       ${col.required ? 'required' : ''}>
            `;
        } else if (col.type === 'number') {
            inputHtml = `
                <input type="number" class="form-control comment-field" 
                       data-column="${escapeHtml(col.name)}" 
                       value="${escapeHtml(savedValue)}"
                       ${col.required ? 'required' : ''}>
            `;
        } else {
            inputHtml = `
                <textarea class="form-control comment-field" 
                          data-column="${escapeHtml(col.name)}" 
                          rows="3"
                          ${col.required ? 'required' : ''}>${escapeHtml(savedValue)}</textarea>
            `;
        }
        
        formGroup.innerHTML = `
            <label>
                ${escapeHtml(col.name)}
                ${col.required ? '<span class="required-indicator">*</span>' : ''}
                ${previousValue && previousValue !== savedValue ? 
                  `<span class="previous-value">Previous: ${escapeHtml(previousValue)}</span>` : ''}
            </label>
            ${inputHtml}
        `;
        
        container.appendChild(formGroup);
    });
    
    container.querySelectorAll('.comment-field').forEach(field => {
        field.addEventListener('input', () => {
            const currentState = {};
            container.querySelectorAll('.comment-field').forEach(f => {
                currentState[f.dataset.column] = f.value;
            });
            saveModalState('commentModal', currentState);
        });
    });
    
    document.getElementById('commentModal').classList.add('show');
}

function closeCommentModal() {
    document.getElementById('commentModal').classList.remove('show');
}

// FIXED: Save comments with proper primary key
async function saveComments() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        
        if (!tabData) return;
        
        const invalidFields = [];
        const commentData = {};
        
        // Collect all field values including the last one
        document.querySelectorAll('#commentFormContainer .comment-field').forEach(field => {
            const column = field.dataset.column;
            const value = field.value;
            const customCol = tab.customColumns.find(c => c.name === column);
            
            if (customCol && customCol.required && !value) {
                invalidFields.push(column);
            }
            
            commentData[column] = value;
        });
        
        if (invalidFields.length > 0) {
            showToast(`Required fields missing: ${invalidFields.join(', ')}`, 'error');
            return;
        }
        
        commentData['UPDATED_BY'] = STATE.currentUser;
        commentData['UPDATED_TIME'] = formatDateTime(new Date());
        
        const selectedData = tabData.filteredData.filter((row, index) => STATE.selectedRows.has(index));
        const customFileName = `${tab.name}_custom_columns.csv`;
        
        if (!tabData.customColumnsData) {
            tabData.customColumnsData = [];
        }
        
        selectedData.forEach(row => {
            const primaryKeyValue = row[tab.primaryKey];
            const existingIndex = tabData.customColumnsData.findIndex(c => c[tab.primaryKey] == primaryKeyValue);
            
            const fullComment = {
                [tab.primaryKey]: primaryKeyValue,
                ...commentData
            };
            
            const oldValue = existingIndex >= 0 ? { ...tabData.customColumnsData[existingIndex] } : null;
            
            if (existingIndex >= 0) {
                tabData.customColumnsData[existingIndex] = { ...tabData.customColumnsData[existingIndex], ...fullComment };
            } else {
                tabData.customColumnsData.push(fullComment);
            }
            
            trackChange('UPDATE_CUSTOM', STATE.currentTab, primaryKeyValue, oldValue, fullComment);
        });
        
        const csvContent = CSV.generate(tabData.customColumnsData);
        await uploadAttachment(customFileName, csvContent);
        
        showToast('Custom fields saved successfully!', 'success');
        clearModalState('commentModal');
        closeCommentModal();
        renderTable();
        
    } catch (error) {
        console.error('Save custom fields error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== EDIT ROW MODAL =====
function openEditRowModal(prefilledData = null, isDuplicate = false) {
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    
    let row;
    let isNewRow = false;
    
    if (prefilledData) {
        row = prefilledData;
        isNewRow = isDuplicate;
    } else if (STATE.selectedRows.size === 1) {
        const selectedIndex = Array.from(STATE.selectedRows)[0];
        row = tabData.filteredData[selectedIndex];
    } else if (STATE.selectedRows.size === 0) {
        row = {};
        isNewRow = true;
    } else {
        showToast('Please select exactly one row to edit', 'warning');
        return;
    }
    
    if (!isNewRow && row[tab.primaryKey]) {
        lockRow(row[tab.primaryKey]);
    }
    
    const allDataColumns = Object.keys(tabData.rawData[0] || {});
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        orderedColumns = [...tab.columnOrder].filter(col => allDataColumns.includes(col));
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = isNewRow ? '-1' : Array.from(STATE.selectedRows)[0];
    container.dataset.primaryKey = row[tab.primaryKey] || '';
    
    const modalState = loadModalState('editRowModal');
    
    // Custom columns section
    if (tab.customColumns && tab.customColumns.length > 0) {
        const customGroup = document.createElement('div');
        customGroup.className = 'form-section';
        customGroup.innerHTML = '<h4>Custom Fields</h4>';
        
        const customRow = tabData.customColumnsData ? 
            tabData.customColumnsData.find(c => c[tab.primaryKey] == row[tab.primaryKey]) : null;
        
        tab.customColumns.forEach(col => {
            const formGroup = document.createElement('div');
            formGroup.className = 'form-group';
            
            const savedValue = modalState[`custom_${col.name}`] || (customRow ? customRow[col.name] : '') || '';
            
            let inputHtml = '';
            
            if (col.type === 'dropdown' && col.options.length > 0) {
                inputHtml = `
                    <select class="form-control" data-column="custom_${escapeHtml(col.name)}" 
                            ${col.required ? 'required' : ''}>
                        <option value="">Select...</option>
                        ${col.options.map(opt => `
                            <option value="${escapeHtml(opt)}" ${savedValue === opt ? 'selected' : ''}>
                                ${escapeHtml(opt)}
                            </option>
                        `).join('')}
                    </select>
                `;
            } else if (col.type === 'date') {
                inputHtml = `
                    <input type="date" class="form-control" 
                           data-column="custom_${escapeHtml(col.name)}" 
                           value="${escapeHtml(savedValue)}"
                           ${col.required ? 'required' : ''}>
                `;
            } else {
                inputHtml = `
                    <input type="text" class="form-control" 
                           data-column="custom_${escapeHtml(col.name)}" 
                           value="${escapeHtml(savedValue)}"
                           ${col.required ? 'required' : ''}>
                `;
            }
            
            formGroup.innerHTML = `
                <label>
                    ${escapeHtml(col.name)}
                    ${col.required ? '<span class="required-indicator">*</span>' : ''}
                </label>
                ${inputHtml}
            `;
            
            customGroup.appendChild(formGroup);
        });
        
        container.appendChild(customGroup);
    }
    
    // Data columns section
    const dataGroup = document.createElement('div');
    dataGroup.className = 'form-section';
    dataGroup.innerHTML = '<h4>Data Fields</h4>';
    
    orderedColumns.forEach(column => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const savedValue = modalState[column] || row[column] || '';
        
        formGroup.innerHTML = `
            <label>${escapeHtml(column)}</label>
            <input type="text" class="form-control" 
                   data-column="${escapeHtml(column)}" 
                   value="${escapeHtml(savedValue)}"
                   ${column === tab.primaryKey && !isNewRow ? 'readonly' : ''}>
        `;
        
        dataGroup.appendChild(formGroup);
    });
    
    container.appendChild(dataGroup);
    
    container.querySelectorAll('.form-control').forEach(field => {
        field.addEventListener('input', () => {
            const currentState = {};
            container.querySelectorAll('.form-control').forEach(f => {
                currentState[f.dataset.column] = f.value;
            });
            saveModalState('editRowModal', currentState);
        });
    });
    
    document.getElementById('editRowTitle').innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            ${isNewRow ? 
                '<line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line>' :
                '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>'}
        </svg>
        ${isNewRow ? 'Add New Row' : 'Edit Row'}
    `;
    
    document.getElementById('editRowModal').classList.add('show');
}

function closeEditRowModal() {
    const container = document.getElementById('editRowFormContainer');
    const primaryKey = container.dataset.primaryKey;
    
    if (primaryKey) {
        unlockRow(primaryKey);
    }
    
    document.getElementById('editRowModal').classList.remove('show');
}

async function saveEditedRow() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        
        const oldRow = rowIndex >= 0 ? tabData.filteredData[rowIndex] : {};
        const newRow = { ...oldRow };
        
        const customData = {};
        let hasCustomData = false;
        
        container.querySelectorAll('.form-control').forEach(input => {
            const column = input.dataset.column;
            
            if (column.startsWith('custom_')) {
                const customColName = column.replace('custom_', '');
                customData[customColName] = input.value;
                hasCustomData = true;
            } else {
                newRow[column] = input.value;
            }
        });
        
        trackChange(rowIndex >= 0 ? 'EDIT' : 'ADD', STATE.currentTab, 
                   newRow[tab.primaryKey], oldRow, newRow);
        
        if (rowIndex >= 0) {
            const primaryKeyValue = oldRow[tab.primaryKey];
            const rawDataIndex = tabData.rawData.findIndex(r => r[tab.primaryKey] === primaryKeyValue);
            
            if (rawDataIndex >= 0) {
                tabData.rawData[rawDataIndex] = newRow;
            }
        } else {
            tabData.rawData.push(newRow);
        }
        
        const csvContent = CSV.generate(tabData.rawData);
        await uploadAttachment(tab.filename, csvContent);
        
        if (hasCustomData && tab.customColumns.length > 0) {
            const customFileName = `${tab.name}_custom_columns.csv`;
            
            if (!tabData.customColumnsData) {
                tabData.customColumnsData = [];
            }
            
            customData[tab.primaryKey] = newRow[tab.primaryKey];
            customData['UPDATED_BY'] = STATE.currentUser;
            customData['UPDATED_TIME'] = formatDateTime(new Date());
            
            const existingIndex = tabData.customColumnsData.findIndex(
                c => c[tab.primaryKey] == newRow[tab.primaryKey]
            );
            
            if (existingIndex >= 0) {
                tabData.customColumnsData[existingIndex] = { 
                    ...tabData.customColumnsData[existingIndex], 
                    ...customData 
                };
            } else {
                tabData.customColumnsData.push(customData);
            }
            
            const customCsv = CSV.generate(tabData.customColumnsData);
            await uploadAttachment(customFileName, customCsv);
        }
        
        await logAudit(rowIndex >= 0 ? 'EDIT' : 'ADD', STATE.currentTab, newRow[tab.primaryKey]);
        
        showToast(rowIndex >= 0 ? 'Row updated!' : 'Row added!', 'success');
        clearModalState('editRowModal');
        closeEditRowModal();
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Save row error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// FIXED: Delete only selected rows, not all
async function deleteSelectedRows() {
    if (STATE.selectedRows.size === 0) {
        showToast('Please select rows to delete', 'warning');
        return;
    }
    
    showConfirm(`Delete ${STATE.selectedRows.size} row(s)?`, async function() {
        try {
            showLoading(true);
            
            const tabData = STATE.allData[STATE.currentTab];
            const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
            
            // Get the primary keys of selected rows
            const selectedData = [];
            STATE.selectedRows.forEach(index => {
                if (tabData.filteredData[index]) {
                    selectedData.push(tabData.filteredData[index]);
                }
            });
            
            const primaryKeysToDelete = selectedData.map(row => row[tab.primaryKey]);
            
            // Track changes
            selectedData.forEach(row => {
                trackChange('DELETE', STATE.currentTab, row[tab.primaryKey], row, null);
            });
            
            // Filter out only the selected rows from rawData
            tabData.rawData = tabData.rawData.filter(r => 
                !primaryKeysToDelete.includes(r[tab.primaryKey])
            );
            
            // Save the updated data
            const csvContent = CSV.generate(tabData.rawData);
            await uploadAttachment(tab.filename, csvContent);
            await logAudit('DELETE', STATE.currentTab, primaryKeysToDelete.join(','));
            
            showToast(`${STATE.selectedRows.size} row(s) deleted!`, 'success');
            
            // Clear selection and reload
            STATE.selectedRows.clear();
            delete STATE.allData[STATE.currentTab];
            await switchTab(STATE.currentTab);
            
        } catch (error) {
            console.error('Delete error:', error);
            showToast('Failed to delete: ' + error.message, 'error');
        } finally {
            showLoading(false);
        }
    });
}

async function logAudit(action, tabName, primaryKey) {
    try {
        let auditData = [];
        const auditFile = `${tabName}_audit_log.csv`;
        
        try {
            const auditText = await fetchAttachment(auditFile);
            auditData = CSV.parse(auditText);
        } catch (error) {
            console.log('Creating new audit log');
        }
        
        auditData.push({
            TIMESTAMP: new Date().toLocaleString(),
            USERNAME: STATE.currentUser,
            ACTION: action,
            TAB: tabName,
            PRIMARY_KEY: primaryKey
        });
        
        const csvContent = CSV.generate(auditData);
        await uploadAttachment(auditFile, csvContent);
    } catch (error) {
        console.warn('Audit log failed:', error);
    }
}

// ===== SELECTION & PAGINATION =====
function selectAllRows() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const pageSize = parseInt(STATE.pageSize);
    const startIndex = (STATE.currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, tabData.filteredData.length);
    
    for (let i = startIndex; i < endIndex; i++) {
        STATE.selectedRows.add(i);
    }
    
    document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = true);
    document.querySelectorAll('tbody tr').forEach(tr => tr.classList.add('selected'));
    
    updateSummary();
    updateButtonVisibility();
}

function unselectAllRows() {
    STATE.selectedRows.clear();
    document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = false);
    document.querySelectorAll('tbody tr').forEach(tr => tr.classList.remove('selected'));
    updateSummary();
    updateButtonVisibility();
}

function toggleRowSelection(index, checked) {
    if (checked) {
        STATE.selectedRows.add(index);
    } else {
        STATE.selectedRows.delete(index);
    }
    
    const row = document.querySelector(`tr[data-index="${index}"]`);
    if (row) {
        row.classList.toggle('selected', checked);
    }
    
    updateSummary();
    updateButtonVisibility();
}

function renderPagination(totalPages, startIndex, endIndex, totalRows) {
    document.getElementById('pageStart').textContent = totalRows > 0 ? startIndex + 1 : 0;
    document.getElementById('pageEnd').textContent = endIndex;
    document.getElementById('totalRowsPage').textContent = totalRows;
    
    const pageNumbersContainer = document.getElementById('pageNumbers');
    pageNumbersContainer.innerHTML = '';
    
    const maxButtons = 5;
    let startPage = Math.max(1, STATE.currentPage - Math.floor(maxButtons / 2));
    let endPage = Math.min(totalPages, startPage + maxButtons - 1);
    
    if (endPage - startPage < maxButtons - 1) {
        startPage = Math.max(1, endPage - maxButtons + 1);
    }
    
    for (let i = startPage; i <= endPage; i++) {
        const btn = document.createElement('button');
        btn.className = 'page-number' + (i === STATE.currentPage ? ' active' : '');
        btn.textContent = i;
        btn.onclick = () => goToPage(i);
        pageNumbersContainer.appendChild(btn);
    }
    
    document.getElementById('prevPageBtn').disabled = STATE.currentPage === 1;
    document.getElementById('nextPageBtn').disabled = STATE.currentPage === totalPages || totalPages === 0;
}

function goToPage(page) {
    STATE.currentPage = page;
    renderTable();
}

// ===== FILTERING =====
function applyFilters() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const customColumnNames = tab && tab.customColumns ? tab.customColumns.map(c => c.name) : [];
    
    let filteredData = [...tabData.rawData];
    
    Object.keys(STATE.activeFilters).forEach(column => {
        const filterValues = STATE.activeFilters[column];
        if (filterValues && filterValues.length > 0) {
            if (customColumnNames.includes(column)) {
                // Filter by custom column
                filteredData = filteredData.filter(row => {
                    const customRow = tabData.customColumnsData ? 
                        tabData.customColumnsData.find(c => c[tab.primaryKey] == row[tab.primaryKey]) : null;
                    const value = customRow ? (customRow[column]?.toString() || '') : '';
                    return filterValues.includes(value);
                });
            } else {
                // Regular column filter
            filteredData = filteredData.filter(row => {
                const value = row[column]?.toString() || '';
                return filterValues.includes(value);
            });
            }
        }
    });
    
    if (STATE.searchQuery) {
        const query = STATE.searchQuery.toLowerCase();
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        
        filteredData = filteredData.filter(row => {
            // Search in regular data columns
            const inDataColumns = Object.values(row).some(value => {
                return value?.toString().toLowerCase().includes(query);
            });
            
            // Search in custom columns if they exist
            let inCustomColumns = false;
            if (tab && tab.customColumns && tab.customColumns.length > 0 && tabData.customColumnsData) {
                const customRow = tabData.customColumnsData.find(c => c[tab.primaryKey] == row[tab.primaryKey]);
                if (customRow) {
                    inCustomColumns = tab.customColumns.some(col => {
                        const value = customRow[col.name];
                        return value?.toString().toLowerCase().includes(query);
                    });
                }
            }
            
            return inDataColumns || inCustomColumns;
        });
    }
    
    tabData.filteredData = filteredData;
    STATE.currentPage = 1;
    
    STATE.tabFilters[STATE.currentTab] = STATE.activeFilters;
    STATE.tabSearch[STATE.currentTab] = STATE.searchQuery;
}

// ===== COLUMN RESIZE FUNCTIONALITY =====
function setupColumnResize() {
    const resizeHandles = document.querySelectorAll('.resize-handle');
    const table = document.getElementById('wsDataTable');
    
    if (!resizeHandles || resizeHandles.length === 0) {
        console.log('No resize handles found');
        return;
    }
    
    console.log(`Setting up ${resizeHandles.length} resize handles`);
    
    resizeHandles.forEach((handle, index) => {
        let isResizing = false;
        let startX = 0;
        let startWidth = 0;
        let currentTh = null;
        
        handle.onmousedown = function(e) {
            console.log('Resize handle mousedown', index);
            e.preventDefault();
            e.stopPropagation();
            
            isResizing = true;
            currentTh = this.parentElement;
            startX = e.pageX;
            startWidth = currentTh.offsetWidth;
            
            this.classList.add('resizing');
            if (table) table.classList.add('resizing');
            
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
        };
        
        document.addEventListener('mousemove', function(e) {
            if (!isResizing || !currentTh) return;
            
            const diff = e.pageX - startX;
            const newWidth = Math.max(100, startWidth + diff);
            
            currentTh.style.width = newWidth + 'px';
            currentTh.style.minWidth = newWidth + 'px';
            currentTh.style.maxWidth = newWidth + 'px';
        });
        
        document.addEventListener('mouseup', function() {
            if (!isResizing) return;
            
            isResizing = false;
            handle.classList.remove('resizing');
            if (table) table.classList.remove('resizing');
            
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            
            currentTh = null;
        });
    });
}

// ===== UTILITIES =====
function formatDateTime(date) {
    // Format: dd/mm/yyyy HH:mm (24-hour format)
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    
    return `${day}/${month}/${year} ${hours}:${minutes}`;
}

function formatNumber(value) {
    const num = parseFloat(value) || 0;
    return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    }).format(num);
}

function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
}

function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (show) {
        overlay.classList.add('show');
    } else {
        overlay.classList.remove('show');
    }
}

function updateButtonVisibility() {
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    if (!tab) return;
    
    const selectedCount = STATE.selectedRows.size;
    
    const editBtn = document.getElementById('editRowButton');
    const addBtn = document.getElementById('addRowButton');
    const deleteBtn = document.getElementById('deleteRowsButton');
    const duplicateBtn = document.getElementById('duplicateRowButton');
    
    // Show buttons based on writeBack setting only (not authorization)
    if (tab.writeBack) {
        editBtn.style.display = 'inline-flex';
        addBtn.style.display = 'inline-flex';
        deleteBtn.style.display = 'inline-flex';
        duplicateBtn.style.display = 'inline-flex';
        
        editBtn.disabled = selectedCount !== 1;
        deleteBtn.disabled = selectedCount === 0;
        duplicateBtn.disabled = selectedCount !== 1;
    } else {
        editBtn.style.display = 'none';
        addBtn.style.display = 'none';
        deleteBtn.style.display = 'none';
        duplicateBtn.style.display = 'none';
    }
    
    const commentsBtn = document.getElementById('addCommentsButton');
    if (tab.customColumns && tab.customColumns.length > 0) {
        commentsBtn.style.display = 'inline-flex';
        commentsBtn.disabled = selectedCount === 0;
    } else {
        commentsBtn.style.display = 'none';
    }
}

function updateSummary() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const filteredData = tabData.filteredData;
    const totalRows = filteredData.length;
    const selectedCount = STATE.selectedRows.size;
    
    document.getElementById('totalRows').textContent = totalRows;
    document.getElementById('selectedRows').textContent = selectedCount;
    
    let sum = 0;
    if (tab.sumColumn) {
        STATE.selectedRows.forEach(index => {
            if (filteredData[index]) {
                const value = parseFloat(filteredData[index][tab.sumColumn]);
                if (!isNaN(value)) {
                    sum += value;
                }
            }
        });
    }
    
    const sumElement = document.getElementById('sumValue');
    sumElement.textContent = formatNumber(sum);
    sumElement.style.color = sum < 0 ? 'var(--danger)' : 'var(--success)';
}

function exportData() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData || tabData.filteredData.length === 0) {
        showToast('No data to export', 'warning');
        return;
    }
    
    const csvContent = CSV.generate(tabData.filteredData);
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${STATE.currentTab}_export_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
    
    showToast('Data exported!', 'success');
}

// ===== EVENT LISTENERS =====
function setupEventListeners() {
    document.getElementById('configButton').addEventListener('click', openConfigModal);
    document.getElementById('saveConfigButton').addEventListener('click', saveConfiguration);
    document.getElementById('addTabButton').addEventListener('click', addTabConfig);
    
    document.getElementById('filterButton').addEventListener('click', openFilterModal);
    document.getElementById('applyFiltersButton').addEventListener('click', applyFiltersFromModal);
    
    document.getElementById('searchInput').addEventListener('input', function() {
        STATE.searchQuery = this.value.toLowerCase();
        applyFilters();
        renderTable();
        updateSummary();
    });
    
    document.getElementById('refreshButton').addEventListener('click', async function() {
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
    });
    
    document.getElementById('exportButton').addEventListener('click', exportData);
    
    document.getElementById('selectAllButton').addEventListener('click', selectAllRows);
    document.getElementById('unselectAllButton').addEventListener('click', unselectAllRows);
    
    document.getElementById('prevPageBtn').addEventListener('click', () => {
        if (STATE.currentPage > 1) {
            goToPage(STATE.currentPage - 1);
        }
    });
    
    document.getElementById('nextPageBtn').addEventListener('click', () => {
        const tabData = STATE.allData[STATE.currentTab];
        if (tabData) {
            const totalPages = Math.ceil(tabData.filteredData.length / STATE.pageSize);
            if (STATE.currentPage < totalPages) {
                goToPage(STATE.currentPage + 1);
            }
        }
    });
    
    document.getElementById('pageSizeSelect').addEventListener('change', function() {
        STATE.pageSize = this.value;
        STATE.currentPage = 1;
        renderTable();
    });
    
    document.getElementById('addCommentsButton').addEventListener('click', openCommentsModal);
    document.getElementById('saveCommentButton').addEventListener('click', saveComments);
    
    document.getElementById('editRowButton').addEventListener('click', () => openEditRowModal());
    document.getElementById('addRowButton').addEventListener('click', () => openEditRowModal(null, false));
    document.getElementById('deleteRowsButton').addEventListener('click', deleteSelectedRows);
    document.getElementById('saveRowButton').addEventListener('click', saveEditedRow);
    document.getElementById('duplicateRowButton').addEventListener('click', () => duplicateRow());
    
    document.getElementById('groupBySelect')?.addEventListener('change', function() {
        groupBy(this.value);
    });
    
    document.getElementById('viewHistoryButton')?.addEventListener('click', openHistoryModal);
    
    window.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal')) {
            e.target.classList.remove('show');
        }
    });
}

// Config modal functions
function openConfigModal() {
    // Auto-detect Page ID if not already set
    if (!CONFIG.pageId) {
        autoDetectConfluenceContext();
    }
    
    document.getElementById('configWorksheetName').value = CONFIG.worksheetName || 'My Worksheet';
    document.getElementById('configBaseUrl').value = CONFIG.confluenceBaseUrl || '';
    document.getElementById('configPageId').value = CONFIG.pageId || '';
    document.getElementById('configTheme').value = CONFIG.theme || 'auto';
    document.getElementById('configAuthorizedUsers').value = CONFIG.authorizedUsers.join(', ');
    
    renderTabsConfig();
    document.getElementById('configModal').classList.add('show');
}

function closeConfigModal() {
    document.getElementById('configModal').classList.remove('show');
}

function renderTabsConfig() {
    const container = document.getElementById('tabsConfigContainer');
    container.innerHTML = '';
    
    STATE.tabs.forEach((tab, index) => {
        const tabConfig = document.createElement('div');
        tabConfig.className = 'tab-config-item';
        tabConfig.innerHTML = `
            <div class="tab-config-header">
                <span class="tab-config-title">Tab ${index + 1}</span>
                <button class="btn btn-danger btn-sm" onclick="removeTabConfig(${index})">×</button>
            </div>
            <div class="form-group">
                <label>Display Name</label>
                <input type="text" class="form-control tab-display-input" value="${escapeHtml(tab.displayName)}" placeholder="Sales Data">
            </div>
            <div class="form-group">
                <label>CSV Filename</label>
                <input type="text" class="form-control tab-filename-input" value="${escapeHtml(tab.filename)}" placeholder="sales_data.csv">
            </div>
            <div class="form-group">
                <label>Upload CSV File</label>
                <input type="file" class="form-control tab-file-upload" accept=".csv" onchange="handleFileUpload(${index}, this.files[0])">
            </div>
            <div class="form-group">
                <label>Primary Key Column</label>
                <input type="text" class="form-control tab-primary-key-input" value="${escapeHtml(tab.primaryKey || 'ID')}" placeholder="ID">
            </div>
            <div class="form-group">
                <label>Sum Column</label>
                <input type="text" class="form-control tab-sum-input" value="${escapeHtml(tab.sumColumn || '')}" placeholder="Amount">
            </div>
            <div class="form-group">
                <label>Column Order (comma-separated)</label>
                <input type="text" class="form-control tab-columns-input" value="${escapeHtml(tab.columnOrder.join(','))}" placeholder="ID,Name,Amount">
            </div>
            <div class="form-group">
                <label>Group By Options (comma-separated)</label>
                <input type="text" class="form-control tab-groupby-input" value="${escapeHtml(tab.groupByOptions ? tab.groupByOptions.join(',') : '')}" placeholder="Category,Status">
            </div>
            <div class="form-group">
                <label>
                    <input type="checkbox" class="tab-writeback-input" ${tab.writeBack ? 'checked' : ''}>
                    Enable Write-Back (Add/Edit/Delete)
                </label>
            </div>
            <div class="custom-columns-section">
                <h5>Custom Columns</h5>
                <div class="custom-columns-container" id="customColumns_${index}">
                    ${tab.customColumns ? tab.customColumns.map((col, colIndex) => {
                        let optionsHtml = '';
                        if (col.type === 'dropdown' && col.options && col.options.length > 0) {
                            optionsHtml = `
                                <div class="color-mapping-container" style="margin-top: 8px; padding: 8px; background: var(--bg-tertiary); border-radius: 4px;">
                                    <small>Color for each option:</small>
                                    ${col.options.map(opt => `
                                        <div style="display: flex; align-items: center; gap: 8px; margin-top: 4px;">
                                            <span style="min-width: 100px;">${escapeHtml(opt)}:</span>
                                            <select class="option-color-select" style="font-size: 12px;">
                                                <option value="default" ${(!col.colorMapping || col.colorMapping[opt] === 'default') ? 'selected' : ''}>Auto</option>
                                                <option value="blue" ${col.colorMapping && col.colorMapping[opt] === 'blue' ? 'selected' : ''}>Blue</option>
                                                <option value="green" ${col.colorMapping && col.colorMapping[opt] === 'green' ? 'selected' : ''}>Green</option>
                                                <option value="yellow" ${col.colorMapping && col.colorMapping[opt] === 'yellow' ? 'selected' : ''}>Yellow</option>
                                                <option value="red" ${col.colorMapping && col.colorMapping[opt] === 'red' ? 'selected' : ''}>Red</option>
                                                <option value="purple" ${col.colorMapping && col.colorMapping[opt] === 'purple' ? 'selected' : ''}>Purple</option>
                                                <option value="pink" ${col.colorMapping && col.colorMapping[opt] === 'pink' ? 'selected' : ''}>Pink</option>
                                                <option value="gray" ${col.colorMapping && col.colorMapping[opt] === 'gray' ? 'selected' : ''}>Gray</option>
                                                <option value="indigo" ${col.colorMapping && col.colorMapping[opt] === 'indigo' ? 'selected' : ''}>Indigo</option>
                                                <option value="teal" ${col.colorMapping && col.colorMapping[opt] === 'teal' ? 'selected' : ''}>Teal</option>
                                                <option value="orange" ${col.colorMapping && col.colorMapping[opt] === 'orange' ? 'selected' : ''}>Orange</option>
                                            </select>
                                        </div>
                                    `).join('')}
                                </div>
                            `;
                        }
                        
                        return `
                            <div class="custom-column-config" style="margin-bottom: 12px; padding: 12px; background: white; border-radius: 4px;">
                                <div style="display: grid; grid-template-columns: 2fr 1fr 2fr auto auto; gap: 8px; align-items: center;">
                            <input type="text" class="custom-col-name" placeholder="Column Name" value="${escapeHtml(col.name)}">
                            <select class="custom-col-type">
                                <option value="text" ${col.type === 'text' ? 'selected' : ''}>Text</option>
                                <option value="dropdown" ${col.type === 'dropdown' ? 'selected' : ''}>Dropdown</option>
                                <option value="date" ${col.type === 'date' ? 'selected' : ''}>Date</option>
                                <option value="number" ${col.type === 'number' ? 'selected' : ''}>Number</option>
                            </select>
                            <input type="text" class="custom-col-options" placeholder="Options (comma-separated)" value="${col.options ? escapeHtml(col.options.join(',')) : ''}">
                            <label>
                                <input type="checkbox" class="custom-col-required" ${col.required ? 'checked' : ''}>
                                Required
                            </label>
                                    <button class="btn btn-danger btn-sm" onclick="this.parentElement.parentElement.remove()">×</button>
                        </div>
                                ${optionsHtml}
                            </div>
                        `;
                    }).join('') : ''}
                </div>
                <button class="btn btn-secondary btn-sm" onclick="addCustomColumn(${index})">Add Custom Column</button>
            </div>
        `;
        
        // Add event listener to show/hide color select based on type
        tabConfig.querySelectorAll('.custom-col-type').forEach(select => {
            select.addEventListener('change', function() {
                const colorSelect = this.parentElement.querySelector('.custom-col-color');
                if (colorSelect) {
                    colorSelect.style.display = this.value === 'dropdown' ? '' : 'none';
                }
            });
        });
        
        container.appendChild(tabConfig);
    });
}

function addCustomColumn(tabIndex) {
    const container = document.getElementById(`customColumns_${tabIndex}`);
    const colConfig = document.createElement('div');
    colConfig.className = 'custom-column-config';
    colConfig.style.cssText = 'margin-bottom: 12px; padding: 12px; background: white; border-radius: 4px;';
    
    colConfig.innerHTML = `
        <div style="display: grid; grid-template-columns: 2fr 1fr 2fr auto auto; gap: 8px; align-items: center;">
        <input type="text" class="custom-col-name" placeholder="Column Name">
        <select class="custom-col-type">
            <option value="text">Text</option>
            <option value="dropdown">Dropdown</option>
            <option value="date">Date</option>
            <option value="number">Number</option>
        </select>
        <input type="text" class="custom-col-options" placeholder="Options (comma-separated)">
        <label>
            <input type="checkbox" class="custom-col-required">
            Required
        </label>
            <button class="btn btn-danger btn-sm" onclick="this.parentElement.parentElement.remove()">×</button>
        </div>
        <div class="color-mapping-container" style="display: none; margin-top: 8px; padding: 8px; background: var(--bg-tertiary); border-radius: 4px;">
            <small>Color for each option:</small>
            <div class="color-options-list"></div>
        </div>
    `;
    
    const typeSelect = colConfig.querySelector('.custom-col-type');
    const optionsInput = colConfig.querySelector('.custom-col-options');
    const colorContainer = colConfig.querySelector('.color-mapping-container');
    const colorOptionsList = colConfig.querySelector('.color-options-list');
    
    // Update color options when dropdown options change
    const updateColorOptions = () => {
        if (typeSelect.value === 'dropdown' && optionsInput.value) {
            const options = optionsInput.value.split(',').map(o => o.trim()).filter(o => o);
            colorContainer.style.display = options.length > 0 ? 'block' : 'none';
            
            colorOptionsList.innerHTML = options.map(opt => `
                <div style="display: flex; align-items: center; gap: 8px; margin-top: 4px;">
                    <span style="min-width: 100px;">${escapeHtml(opt)}:</span>
                    <select class="option-color-select" style="font-size: 12px;">
                        <option value="default">Auto</option>
                        <option value="blue">Blue</option>
                        <option value="green">Green</option>
                        <option value="yellow">Yellow</option>
                        <option value="red">Red</option>
                        <option value="purple">Purple</option>
                        <option value="pink">Pink</option>
                        <option value="gray">Gray</option>
                        <option value="indigo">Indigo</option>
                        <option value="teal">Teal</option>
                        <option value="orange">Orange</option>
                    </select>
                </div>
            `).join('');
        } else {
            colorContainer.style.display = 'none';
        }
    };
    
    typeSelect.addEventListener('change', updateColorOptions);
    optionsInput.addEventListener('input', updateColorOptions);
    
    container.appendChild(colConfig);
}

function addTabConfig() {
    if (STATE.tabs.length >= 5) {
        showToast('Maximum 5 tabs allowed', 'warning');
        return;
    }
    
    STATE.tabs.push({
        key: `TAB_${STATE.tabs.length + 1}`,
        name: '',
        displayName: 'New Tab',
        filename: '',
        primaryKey: 'ID',
        sumColumn: '',
        writeBack: false,
        customColumns: [],
        columnOrder: [],
        groupByOptions: []
    });
    renderTabsConfig();
}

function removeTabConfig(index) {
    if (STATE.tabs.length === 1) {
        showToast('Cannot remove the last tab', 'warning');
        return;
    }
    STATE.tabs.splice(index, 1);
    renderTabsConfig();
}

// Filter modal functions
function openFilterModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.rawData;
    if (data.length === 0) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    
    const allDataColumns = Object.keys(data[0]);
    const customColumnNames = tab.customColumns ? tab.customColumns.map(c => c.name) : [];
    const allColumns = [...customColumnNames, ...allDataColumns];
    
    const container = document.getElementById('filterColumnsContainer');
    container.innerHTML = '';
    
    allColumns.forEach(column => {
        const isCustom = customColumnNames.includes(column);
        let uniqueValues = [];
        
        if (isCustom) {
            if (tabData.customColumnsData) {
                uniqueValues = [...new Set(tabData.customColumnsData.map(row => row[column]?.toString() || ''))].sort();
            }
        } else {
            uniqueValues = [...new Set(data.map(row => row[column]?.toString() || ''))].sort();
        }
        
        const isActive = STATE.activeFilters[column] && STATE.activeFilters[column].length > 0;
        
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-column-group' + (isActive ? ' active' : '');
        
        filterGroup.innerHTML = `
            <label>${escapeHtml(column)}</label>
            <select multiple class="filter-select" data-column="${escapeHtml(column)}">
                ${uniqueValues.map(val => {
                    const isSelected = STATE.activeFilters[column]?.includes(val);
                    return `<option value="${escapeHtml(val)}" ${isSelected ? 'selected' : ''}>${escapeHtml(val)}</option>`;
                }).join('')}
            </select>
        `;
        container.appendChild(filterGroup);
    });
    
    updateFilterStatus();
    document.getElementById('filterModal').classList.add('show');
}

function closeFilterModal() {
    document.getElementById('filterModal').classList.remove('show');
}

function applyFiltersFromModal() {
    STATE.activeFilters = {};
    
    document.querySelectorAll('.filter-select').forEach(select => {
        const column = select.dataset.column;
        const selectedOptions = Array.from(select.selectedOptions).map(opt => opt.value);
        
        if (selectedOptions.length > 0) {
            STATE.activeFilters[column] = selectedOptions;
        }
    });
    
    const activeCount = Object.keys(STATE.activeFilters).length;
    const badge = document.getElementById('filterBadge');
    if (activeCount > 0) {
        badge.textContent = activeCount;
        badge.style.display = 'inline-flex';
    } else {
        badge.style.display = 'none';
    }
    
    applyFilters();
    renderTable();
    updateSummary();
    closeFilterModal();
}

function clearFilters() {
    STATE.activeFilters = {};
    STATE.searchQuery = '';
    document.getElementById('searchInput').value = '';
    document.getElementById('filterBadge').style.display = 'none';
    
    STATE.tabFilters[STATE.currentTab] = {};
    STATE.tabSearch[STATE.currentTab] = '';
    
    applyFilters();
    renderTable();
    updateSummary();
    closeFilterModal();
}

function updateFilterStatus() {
    const count = Object.keys(STATE.activeFilters).length;
    const status = document.getElementById('filterStatus');
    if (count > 0) {
        status.textContent = `${count} filter${count > 1 ? 's' : ''} active`;
        status.style.color = 'var(--primary)';
    } else {
        status.textContent = 'No filters active';
        status.style.color = 'var(--text-muted)';
    }
}

// History modal
async function openHistoryModal() {
    const modal = document.getElementById('historyModal');
    if (!modal) return;
    
    const container = document.getElementById('historyContainer');
    container.innerHTML = '<p>Loading history...</p>';
    
    // Show modal first
    modal.classList.add('show');
    
    try {
        // Try to load history from file
        if (STATE.currentTab) {
            const historyFile = `${STATE.currentTab}_change_history.csv`;
            try {
                const historyText = await fetchAttachment(historyFile);
                const historyData = CSV.parse(historyText);
                if (historyData && historyData.length > 0) {
                    STATE.changeHistory = historyData;
                }
            } catch (error) {
                console.log('No existing history file, using in-memory history');
            }
        }
        
        // Get last 100 entries
        const recentHistory = STATE.changeHistory.slice(-100).reverse();
        
        // Clear loading message
        container.innerHTML = '';
        
        if (recentHistory.length === 0) {
            container.innerHTML = '<p style="padding: 20px; text-align: center;">No change history available</p>';
        } else {
            const table = document.createElement('table');
            table.className = 'history-table';
            table.innerHTML = `
                <thead>
                    <tr>
                        <th>Time</th>
                        <th>User</th>
                        <th>Action</th>
                        <th>Tab</th>
                        <th>Primary Key</th>
                    </tr>
                </thead>
                <tbody>
                    ${recentHistory.map(change => `
                        <tr>
                            <td>${escapeHtml(new Date(change.timestamp || change.TIMESTAMP).toLocaleString())}</td>
                            <td>${escapeHtml(change.user || change.USERNAME)}</td>
                            <td>${escapeHtml(change.action || change.ACTION)}</td>
                            <td>${escapeHtml(change.tab || change.TAB)}</td>
                            <td>${escapeHtml(change.primaryKey || change.PRIMARY_KEY)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            `;
            container.appendChild(table);
        }
    } catch (error) {
        console.error('Error loading history:', error);
        container.innerHTML = '<p style="padding: 20px; text-align: center; color: var(--danger);">Error loading history</p>';
    }
}

function closeHistoryModal() {
    const modal = document.getElementById('historyModal');
    if (modal) modal.classList.remove('show');
}

// Global functions
window.closeConfigModal = closeConfigModal;
window.closeFilterModal = closeFilterModal;
window.clearFilters = clearFilters;
window.closeCommentModal = closeCommentModal;
window.closeEditRowModal = closeEditRowModal;
window.closeConfirmModal = closeConfirmModal;
window.closeHistoryModal = closeHistoryModal;
window.closePresenceModal = closePresenceModal;
window.removeTabConfig = removeTabConfig;
window.addCustomColumn = addCustomColumn;
window.handleFileUpload = handleFileUpload;
window.duplicateRow = duplicateRow;
window.toggleGroup = toggleGroup;

console.log('Worksheet Tool Optimized v6 loaded!');
