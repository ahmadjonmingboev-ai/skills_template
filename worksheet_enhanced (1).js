// Enhanced Worksheet Tool - Complete Implementation with All Features and Fixes
// Part 1: Core Utilities and Configuration

// ===== CSV UTILITIES =====
const CSV = {
    parse: function(csvText) {
        if (!csvText) return [];
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
        const str = String(field || '');
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
    }
};

// ===== CONFIGURATION =====
const CONFIG = {
    confluenceBaseUrl: '',
    pageId: '',
    worksheetName: 'Worksheet Tool',
    theme: 'light',
    permissions: {
        admins: [],
        editors: [],
        viewers: []
    },
    configFile: 'worksheet_config.csv',
    historyFile: 'worksheet_history.csv',
    locksFile: 'worksheet_locks.csv'
};

// ===== STATE MANAGEMENT =====
const STATE = {
    currentUser: null,
    userEmail: null,
    userPermission: null,
    tabs: [],
    currentTab: null,
    tabData: {},
    selectedRows: new Set(),
    currentPage: 1,
    pageSize: 25,
    sortColumn: null,
    sortDirection: 'asc',
    groupColumn: null,
    groupCollapsed: false,
    filters: [],
    searchQuery: '',
    modalData: {},  // FIX 3: Store modal form data
    rowLocks: {},
    changeHistory: [],
    contextRow: null
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', async () => {
    try {
        showLoading('Initializing...');
        await getCurrentUser();
        await loadConfiguration();
        await loadTabs();
        setupEventListeners();
        initializeUI();
        
        if (STATE.tabs.length > 0) {
            await switchTab(STATE.tabs[0].id);
        } else {
            showEmptyState();
        }
        
        hideLoading();
    } catch (error) {
        console.error('Initialization error:', error);
        showToast('Failed to initialize: ' + error.message, 'error');
        hideLoading();
    }
});

// ===== USER MANAGEMENT =====
async function getCurrentUser() {
    try {
        const response = await fetch(`${CONFIG.confluenceBaseUrl}/rest/api/user/current`, {
            credentials: 'include'
        });
        
        if (response.ok) {
            const user = await response.json();
            STATE.currentUser = user.displayName || user.username || 'Unknown User';
            STATE.userEmail = user.email || '';
            
            // Determine user permission level
            if (CONFIG.permissions.admins.includes(STATE.userEmail)) {
                STATE.userPermission = 'admin';
            } else if (CONFIG.permissions.editors.includes(STATE.userEmail)) {
                STATE.userPermission = 'editor';
            } else if (CONFIG.permissions.viewers.includes(STATE.userEmail)) {
                STATE.userPermission = 'viewer';
            } else {
                STATE.userPermission = 'viewer';
            }
            
            // Update UI
            document.getElementById('userName').textContent = STATE.currentUser;
            const initials = STATE.currentUser.split(' ').map(n => n[0]).join('').toUpperCase();
            document.getElementById('userInitials').textContent = initials.substring(0, 2);
        }
    } catch (error) {
        console.error('Failed to get current user:', error);
        STATE.currentUser = 'Guest';
        STATE.userPermission = 'viewer';
    }
}

// ===== CONFIGURATION LOADING =====
async function loadConfiguration() {
    try {
        const configData = await fetchAttachment(CONFIG.configFile);
        const config = CSV.parse(configData);
        
        config.forEach(row => {
            if (row.key === 'worksheetName') CONFIG.worksheetName = row.value;
            if (row.key === 'confluenceBaseUrl') CONFIG.confluenceBaseUrl = row.value;
            if (row.key === 'pageId') CONFIG.pageId = row.value;
            if (row.key === 'theme') CONFIG.theme = row.value;
            if (row.key === 'admins') CONFIG.permissions.admins = row.value.split(',').map(e => e.trim());
            if (row.key === 'editors') CONFIG.permissions.editors = row.value.split(',').map(e => e.trim());
            if (row.key === 'viewers') CONFIG.permissions.viewers = row.value.split(',').map(e => e.trim());
        });
        
        document.getElementById('worksheetName').textContent = CONFIG.worksheetName;
        applyTheme(CONFIG.theme);
        
    } catch (error) {
        console.log('No configuration file found, using defaults');
        setTimeout(() => openModal('configModal'), 500);
    }
}

async function loadTabs() {
    try {
        const tabsData = await fetchAttachment('worksheet_tabs.csv');
        const tabs = CSV.parse(tabsData);
        
        STATE.tabs = tabs.map(tab => ({
            id: tab.id || generateId(),
            name: tab.name,
            fileName: tab.fileName,
            primaryKey: tab.primaryKey || 'ID',
            sumColumn: tab.sumColumn || null,
            writeBack: tab.writeBack === 'true',
            grouping: tab.grouping === 'true',
            locking: tab.locking === 'true',
            customColumns: JSON.parse(tab.customColumns || '[]'),
            customFileName: `${tab.id}_custom_columns.csv`
        }));
        
    } catch (error) {
        console.log('No tabs configuration found');
        STATE.tabs = [];
    }
}

// ===== ATTACHMENT OPERATIONS =====
async function fetchAttachment(filename) {
    const response = await fetch(
        `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`,
        { credentials: 'include' }
    );
    
    const data = await response.json();
    const attachment = data.results.find(a => a.title === filename);
    
    if (!attachment) {
        throw new Error(`Attachment ${filename} not found`);
    }
    
    const content = await fetch(CONFIG.confluenceBaseUrl + attachment._links.download, {
        credentials: 'include'
    });
    
    return await content.text();
}

async function uploadAttachment(filename, content) {
    const attachmentsResponse = await fetch(
        `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`,
        { credentials: 'include' }
    );
    
    const attachments = await attachmentsResponse.json();
    const existing = attachments.results.find(a => a.title === filename);
    
    const blob = new Blob([content], { type: 'text/csv' });
    const formData = new FormData();
    formData.append('file', blob, filename);
    formData.append('minorEdit', 'true');
    
    let url = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    if (existing) {
        url += `/${existing.id}/data`;
    }
    
    const response = await fetch(u
// ===== COMMENTS HANDLING =====
async function openCommentsModal() {
    const selectedCount = STATE.selectedRows.size;
    if (selectedCount === 0) {
        showToast('Please select rows to add comments', 'warning');
        return;
    }
    
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab].filteredData;
    const selectedIndices = Array.from(STATE.selectedRows);
    
    document.getElementById('commentRowCount').textContent = selectedCount;
    
    const container = document.getElementById('commentsFormContainer');
    container.innerHTML = '';
    
    // FIX 4: Show previous values for comments
    const firstRow = data[selectedIndices[0]];
    
    // Create fields for custom columns
    tab.customColumns.forEach(col => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const existingValue = firstRow[col.name] || STATE.modalData[col.name] || '';
        
        if (col.type === 'dropdown') {
            formGroup.innerHTML = `
                <label>${col.name} ${col.required ? '<span class="text-danger">*</span>' : ''}</label>
                <select class="form-control" data-column="${col.name}" ${col.required ? 'required' : ''}>
                    <option value="">-- Select --</option>
                    ${col.options.map(opt => `
                        <option value="${opt}" ${existingValue === opt ? 'selected' : ''}>${opt}</option>
                    `).join('')}
                </select>
            `;
        } else if (col.type === 'user') {
            formGroup.innerHTML = `
                <label>${col.name} ${col.required ? '<span class="text-danger">*</span>' : ''}</label>
                <input type="text" class="form-control" data-column="${col.name}" 
                       value="${existingValue || STATE.currentUser}" ${col.required ? 'required' : ''}>
            `;
        } else {
            formGroup.innerHTML = `
                <label>${col.name} ${col.required ? '<span class="text-danger">*</span>' : ''}</label>
                <textarea class="form-control" data-column="${col.name}" 
                          rows="3" ${col.required ? 'required' : ''}>${existingValue}</textarea>
            `;
        }
        
        container.appendChild(formGroup);
    });
    
    // Add info about auto-filled fields
    const infoBox = document.createElement('div');
    infoBox.className = 'info-box mt-2';
    infoBox.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="10"></circle>
            <line x1="12" y1="16" x2="12" y2="12"></line>
            <line x1="12" y1="8" x2="12.01" y2="8"></line>
        </svg>
        <span>System fields: USERNAME (${STATE.currentUser}), UPDATED_TIME (Current time)</span>
    `;
    container.appendChild(infoBox);
    
    // Store form data
    document.querySelectorAll('#commentsFormContainer .form-control').forEach(field => {
        field.addEventListener('input', () => {
            STATE.modalData[field.dataset.column] = field.value;
        });
    });
    
    openModal('commentsModal');
}

async function saveComments() {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab];
    const selectedIndices = Array.from(STATE.selectedRows);
    
    // Validate required fields
    const formData = {};
    let isValid = true;
    
    document.querySelectorAll('#commentsFormContainer .form-control').forEach(field => {
        const column = field.dataset.column;
        const value = field.value;
        
        if (field.hasAttribute('required') && !value) {
            field.classList.add('error');
            isValid = false;
        } else {
            field.classList.remove('error');
            formData[column] = value;
        }
    });
    
    if (!isValid) {
        showToast('Please fill all required fields', 'warning');
        return;
    }
    
    // Add system fields
    formData['USERNAME'] = STATE.currentUser;
    formData['UPDATED_TIME'] = new Date().toISOString();
    
    // Update selected rows
    for (const index of selectedIndices) {
        const row = data.filteredData[index];
        const oldValues = {};
        
        // Store old values for history
        Object.keys(formData).forEach(key => {
            if (key !== 'USERNAME' && key !== 'UPDATED_TIME') {
                oldValues[key] = row[key];
                row[key] = formData[key];
            }
        });
        
        // Log changes
        await logChange('COMMENT', tab.id, row[tab.primaryKey], oldValues, formData);
        
        // Update original data
        const originalIndex = data.originalData.findIndex(r => 
            r[tab.primaryKey] === row[tab.primaryKey]
        );
        
        if (originalIndex !== -1) {
            Object.assign(data.originalData[originalIndex], formData);
        }
    }
    
    // FIX 1: Save all columns including the last one
    await saveCustomColumns();
    
    closeModal('commentsModal');
    showToast('Comments saved successfully', 'success');
    renderTable();
}

// ===== SAVE OPERATIONS =====
async function saveTabData() {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab];
    
    // Separate main data and custom columns
    const mainData = data.originalData.map(row => {
        const mainRow = {};
        
        // Only include non-custom columns in main data
        Object.keys(row).forEach(key => {
            if (!tab.customColumns.some(c => c.name === key)) {
                mainRow[key] = row[key];
            }
        });
        
        return mainRow;
    });
    
    // Save main data
    const csvContent = CSV.generate(mainData);
    await uploadAttachment(tab.fileName, csvContent);
    
    // Update last saved time
    updateLastSaved();
}

async function saveCustomColumns() {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab];
    
    if (tab.customColumns.length === 0) return;
    
    // FIX 1: Ensure all custom columns are saved including the last one
    const customData = data.originalData.map(row => {
        const customRow = {
            [tab.primaryKey]: row[tab.primaryKey],
            'USERNAME': STATE.currentUser,
            'UPDATED_TIME': new Date().toISOString()
        };
        
        // Include ALL custom columns (FIX for last column not saving)
        tab.customColumns.forEach((col, index) => {
            customRow[col.name] = row[col.name] || '';
        });
        
        return customRow;
    }).filter(row => {
        // Only include rows that have custom column data
        return tab.customColumns.some(col => row[col.name]);
    });
    
    if (customData.length > 0) {
        const csvContent = CSV.generate(customData);
        await uploadAttachment(tab.customFileName, csvContent);
    }
}

// ===== MODALS =====
function openModal(modalId) {
    document.getElementById(modalId).classList.add('show');
    document.body.style.overflow = 'hidden';
}

function closeModal(modalId) {
    document.getElementById(modalId).classList.remove('show');
    document.body.style.overflow = '';
}

function openEditModal(rowIndex, title) {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab].filteredData;
    
    // Use stored modal data or get from row
    const row = rowIndex !== null ? data[rowIndex] : (STATE.modalData.editRow || {});
    
    document.getElementById('editModalTitle').textContent = title;
    
    const container = document.getElementById('editFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = rowIndex;
    
    // Get all columns
    const columns = data.length > 0 ? Object.keys(data[0] || {}) : 
                   Object.keys(STATE.tabData[STATE.currentTab].originalData[0] || {});
    
    columns.forEach(col => {
        if (col.startsWith('_')) return; // Skip internal columns
        
        const customCol = tab.customColumns.find(c => c.name === col);
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const value = row[col] || STATE.modalData[col] || '';
        const isPrimaryKey = col === tab.primaryKey;
        
        if (customCol) {
            if (customCol.type === 'dropdown') {
                formGroup.innerHTML = `
                    <label>${col} ${customCol.required ? '<span class="text-danger">*</span>' : ''}</label>
                    <select class="form-control" data-column="${col}" ${customCol.required ? 'required' : ''}>
                        <option value="">-- Select --</option>
                        ${customCol.options.map(opt => `
                            <option value="${opt}" ${value === opt ? 'selected' : ''}>${opt}</option>
                        `).join('')}
                    </select>
                `;
            } else if (customCol.type === 'user') {
                formGroup.innerHTML = `
                    <label>${col} ${customCol.required ? '<span class="text-danger">*</span>' : ''}</label>
                    <input type="text" class="form-control" data-column="${col}" 
                           value="${value}" ${customCol.required ? 'required' : ''}>
                `;
            } else {
                formGroup.innerHTML = `
                    <label>${col} ${customCol.required ? '<span class="text-danger">*</span>' : ''}</label>
                    <input type="text" class="form-control" data-column="${col}" 
                           value="${value}" ${customCol.required ? 'required' : ''}>
                `;
            }
        } else {
            formGroup.innerHTML = `
                <label>${col}</label>
                <input type="text" class="form-control" data-column="${col}" 
                       value="${value}" ${isPrimaryKey && rowIndex !== null ? 'disabled' : ''}>
            `;
        }
        
        container.appendChild(formGroup);
        
        // Store form data on change
        const field = formGroup.querySelector('.form-control');
        if (field) {
            field.addEventListener('input', () => {
                STATE.modalData[col] = field.value;
            });
        }
    });
    
    openModal('editModal');
}

async function saveEditRow() {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab];
    const container = document.getElementById('editFormContainer');
    const rowIndex = container.dataset.rowIndex !== 'null' ? 
                     parseInt(container.dataset.rowIndex) : null;
    
    // Collect form data
    const formData = {};
    let isValid = true;
    
    document.querySelectorAll('#editFormContainer .form-control').forEach(field => {
        const column = field.dataset.column;
        const value = field.value;
        
        if (field.hasAttribute('required') && !value) {
            field.classList.add('error');
            isValid = false;
        } else {
            field.classList.remove('error');
            formData[column] = value;
        }
    });
    
    if (!isValid) {
        showToast('Please fill all required fields', 'warning');
        return;
    }
    
    if (rowIndex !== null) {
        // Edit existing row
        const oldRow = { ...data.filteredData[rowIndex] };
        
        // Update filtered data
        Object.assign(data.filteredData[rowIndex], formData);
        
        // Update original data
        const originalIndex = data.originalData.findIndex(r => 
            r[tab.primaryKey] === oldRow[tab.primaryKey]
        );
        
        if (originalIndex !== -1) {
            Object.assign(data.originalData[originalIndex], formData);
        }
        
        await logChange('EDIT', tab.id, oldRow[tab.primaryKey], oldRow, formData);
        
    } else {
        // Add new row
        if (!formData[tab.primaryKey]) {
            formData[tab.primaryKey] = generateId();
        }
        
        data.originalData.push(formData);
        data.filteredData.push(formData);
        
        await logChange('ADD', tab.id, formData[tab.primaryKey], null, formData);
    }
    
    await saveTabData();
    await saveCustomColumns();
    
    closeModal('editModal');
    showToast('Changes saved successfully', 'success');
    renderTable();
    updateButtonStates();
}

// ===== HISTORY =====
async function openHistoryModal() {
    await loadHistory();
    openModal('historyModal');
}

async function loadHistory() {
    try {
        const historyData = await fetchAttachment(CONFIG.historyFile);
        const history = CSV.parse(historyData);
        
        // Filter by date and user if specified
        let filtered = history;
        
        const dateFrom = document.getElementById('historyDateFrom').value;
        const dateTo = document.getElementById('historyDateTo').value;
        const user = document.getElementById('historyUser').value;
        
        if (dateFrom) {
            filtered = filtered.filter(h => new Date(h.timestamp) >= new Date(dateFrom));
        }
        
        if (dateTo) {
            filtered = filtered.filter(h => new Date(h.timestamp) <= new Date(dateTo));
        }
        
        if (user) {
            filtered = filtered.filter(h => h.user === user);
        }
        
        // Sort by timestamp descending
        filtered.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        
        // Render history
        const container = document.getElementById('historyContainer');
        container.innerHTML = filtered.map(item => `
            <div class="history-item">
                <div class="history-details">
                    <span class="history-user">${item.user}</span>
                    <span class="history-action">${item.action}</span>
                    <span class="history-time">${new Date(item.timestamp).toLocaleString()}</span>
                    <div class="history-changes">${item.changes || 'No details'}</div>
                </div>
            </div>
        `).join('');
        
    } catch (error) {
        console.log('No history available');
        document.getElementById('historyContainer').innerHTML = 
            '<p class="text-muted text-center">No history available</p>';
    }
}

async function logChange(action, tabId, rowId, oldData, newData) {
    try {
        let history = [];
        try {
            const historyData = await fetchAttachment(CONFIG.historyFile);
            history = CSV.parse(historyData);
        } catch (e) {}
        
        const changes = [];
        if (oldData && newData) {
            Object.keys(newData).forEach(key => {
                if (oldData[key] !== newData[key]) {
                    changes.push(`${key}: ${oldData[key]} → ${newData[key]}`);
                }
            });
        }
        
        history.push({
            timestamp: new Date().toISOString(),
            user: STATE.currentUser,
            action: action,
            tabId: tabId,
            rowId: rowId,
            changes: changes.join(', ')
        });
        
        const csvContent = CSV.generate(history);
        await uploadAttachment(CONFIG.historyFile, csvContent);
        
    } catch (error) {
        console.log('Failed to log history:', error);
    }
}

// ===== UI HELPERS =====
function renderTabs() {
    const container = document.getElementById('tabsContainer');
    
    container.innerHTML = STATE.tabs.map(tab => `
        <button class="tab ${tab.id === STATE.currentTab ? 'active' : ''}" 
                onclick="switchTab('${tab.id}')">
            ${tab.name}
        </button>
    `).join('');
}

function updateSummaryCards() {
    const data = STATE.tabData[STATE.currentTab];
    if (!data) return;
    
    document.getElementById('totalRecords').textContent = data.originalData.length;
    
    // Update other cards based on your business logic
    document.getElementById('activeItems').textContent = 
        data.originalData.filter(row => row.Status === 'Active').length;
    
    document.getElementById('pendingReview').textContent = 
        data.originalData.filter(row => row.Status === 'Pending').length;
    
    document.getElementById('dueToday').textContent = 
        data.originalData.filter(row => {
            const dueDate = new Date(row.DueDate);
            const today = new Date();
            return dueDate.toDateString() === today.toDateString();
        }).length;
}

function updateButtonStates() {
    const hasSelection = STATE.selectedRows.size > 0;
    const singleSelection = STATE.selectedRows.size === 1;
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    
    // Check permissions
    const canEdit = STATE.userPermission === 'admin' || STATE.userPermission === 'editor';
    const canDelete = STATE.userPermission === 'admin';
    
    // Update button states
    document.getElementById('editRowBtn').disabled = !singleSelection || !canEdit || !tab?.writeBack;
    document.getElementById('duplicateRowBtn').disabled = !singleSelection || !canEdit || !tab?.writeBack;
    document.getElementById('deleteRowBtn').disabled = !hasSelection || !canDelete || !tab?.writeBack;
    document.getElementById('bulkEditBtn').disabled = !hasSelection || !canEdit || !tab?.writeBack;
    document.getElementById('commentBtn').disabled = !hasSelection || !canEdit;
}

function updateStatusBar() {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab];
    
    if (!data) return;
    
    // Update selected count
    document.getElementById('selectedCount').textContent = 
        `${STATE.selectedRows.size} selected`;
    
    // Calculate sum if sum column is configured
    if (tab?.sumColumn) {
        let sum = 0;
        STATE.selectedRows.forEach(index => {
            const row = data.filteredData[index];
            if (row) {
                const value = parseFloat(row[tab.sumColumn]);
                if (!isNaN(value)) {
                    sum += value;
                }
            }
        });
        
        // FIX: Show negative sums in red
        const sumElement = document.getElementById('sumValue');
        sumElement.textContent = `Sum: ${formatNumber(sum)}`;
        sumElement.style.color = sum < 0 ? 'var(--danger)' : '';
    } else {
        document.getElementById('sumValue').textContent = 'Sum: N/A';
    }
}

function updateLastSaved() {
    const now = new Date();
    document.getElementById('lastSaved').textContent = 
        `Last saved: ${now.toLocaleTimeString()}`;
}

function showEmptyState() {
    document.getElementById('tableContainer').innerHTML = `
        <div class="empty-state">
            <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
                <path d="M9 9a3 3 0 1 1 6 0"></path>
                <path d="M12 12v3"></path>
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="17" x2="12.01" y2="17"></line>
            </svg>
            <h3>No tabs configured</h3>
            <p>Add a new tab to get started</p>
            <button class="btn-primary" onclick="addNewTab()">Add Tab</button>
        </div>
    `;
}

function showLoading(message = 'Loading...') {
    const overlay = document.getElementById('loadingOverlay');
    overlay.querySelector('p').textContent = message;
    overlay.classList.add('show');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('show');
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icons = {
        success: '<svg class="toast-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>',
        error: '<svg class="toast-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
        warning: '<svg class="toast-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>',
        info: '<svg class="toast-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>'
    };
    
    toast.innerHTML = `
        ${icons[type] || icons.info}
        <div class="toast-message">${message}</div>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 5000);
}

function applyTheme(theme) {
    if (theme === 'auto') {
        const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        document.documentElement.setAttribute('data-theme', prefersDark ? 'dark' : 'light');
    } else {
        document.documentElement.setAttribute('data-theme', theme);
    }
}

// ===== UTILITY FUNCTIONS =====
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
}

function formatNumber(value) {
    return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    }).format(value);
}

function formatFileSize(bytes) {
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    if (bytes === 0) return '0 Bytes';
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, m => map[m]);
}

// ===== EVENT LISTENERS =====
function setupEventListeners() {
    // Search
    document.getElementById('globalSearch').addEventListener('input', (e) => {
        STATE.searchQuery = e.target.value;
        STATE.currentPage = 1;
        renderTable();
    });
    
    // Page size
    document.getElementById('pageSize').addEventListener('change', (e) => {
        STATE.pageSize = parseInt(e.target.value);
        STATE.currentPage = 1;
        renderTable();
    });
    
    // Toolbar buttons
    document.getElementById('addRowBtn').addEventListener('click', addRow);
    document.getElementById('editRowBtn').addEventListener('click', editRow);
    document.getElementById('duplicateRowBtn').addEventListener('click', duplicateRow);
    document.getElementById('deleteRowBtn').addEventListener('click', deleteRows);
    document.getElementById('bulkEditBtn').addEventListener('click', bulkEdit);
    document.getElementById('commentBtn').addEventListener('click', openCommentsModal);
    
    // History
    document.getElementById('historyBtn').addEventListener('click', openHistoryModal);
    
    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        // Escape to close modals
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal.show').forEach(modal => {
                closeModal(modal.id);
            });
        }
    });
}

function initializeUI() {
    // Initialize tooltips, connection status, etc.
}

// ===== EXPORT GLOBAL FUNCTIONS FOR HTML =====
window.sortTable = sortTable;
window.toggleSelectAll = toggleSelectAll;
window.toggleRowSelection = toggleRowSelection;
window.changePage = changePage;
window.showContextMenu = function() {};
window.closeModal = closeModal;
window.openModal = openModal;
window.saveEditRow = saveEditRow;
window.saveComments = saveComments;
window.switchTab = switchTab;
window.contextEdit = function() {};
window.contextDuplicate = function() {};
window.contextComment = function() {};
window.contextDelete = function() {};
window.toggleGroup = function() {};
window.loadHistory = loadHistory;

console.log('Enhanced Worksheet Tool initialized');
