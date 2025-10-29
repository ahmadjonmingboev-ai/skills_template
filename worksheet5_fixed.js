// WORKSHEET TOOL - Final Version with All Fixes
// ================================================

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
        if (data.length === 0) return '';
        
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
        <div class="toast-message">${message}</div>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.classList.add('removing');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ===== CONFIRM MODAL =====
function showConfirm(message, onConfirm) {
    document.getElementById('confirmMessage').textContent = message;
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
    commentDBFile: 'commentdb.csv',
    auditLogFile: 'audit_log.csv',
    worksheetName: 'Worksheet',
    primaryKey: 'ITEM_ID',
    theme: 'auto'
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
    commentColumns: [],
    sortColumn: null,        // FIX 3: Add sorting state
    sortDirection: 'asc'     // FIX 3: Add sorting state
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', async function() {
    console.log('Worksheet Tool (Final Version) initializing...');
    await init();
});

async function init() {
    try {
        showLoading(true);
        await getCurrentUser();
        
        try {
            await loadConfiguration();
        } catch (error) {
            console.warn('Config not found');
            showLoading(false);
            setTimeout(() => openConfigModal(), 500);
            return;
        }
        
        applyTheme(CONFIG.theme);
        document.getElementById('worksheetName').textContent = CONFIG.worksheetName;
        setupEventListeners();
        
        if (STATE.tabs.length > 0) {
            await switchTab(STATE.tabs[0].name);
        }
        
        showLoading(false);
    } catch (error) {
        console.error('Init error:', error);
        showToast('Failed to initialize: ' + error.message, 'error');
        showLoading(false);
    }
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
        } else {
            STATE.currentUser = 'Guest';
        }
    } catch (error) {
        STATE.currentUser = 'Guest';
    }
    document.getElementById('currentUser').textContent = STATE.currentUser;
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

// ===== CONFIGURATION =====
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
        } else if (field === 'PRIMARY_KEY') {
            CONFIG.primaryKey = value;
        } else if (field === 'THEME') {
            CONFIG.theme = value;
        } else if (field.match(/^TAB_FILE_\d+$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            
            if (!STATE.tabs.find(t => t.key === tabKey)) {
                STATE.tabs.push({
                    key: tabKey,
                    name: value.replace('.csv', ''),
                    displayName: value,
                    filename: '',
                    sumColumn: '',
                    writeBack: false,
                    commentsEnabled: true,
                    columnOrder: []
                });
            }
        } else if (field.match(/^TAB_FILE_\d+_NAME$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.filename = value;
        } else if (field.match(/^TAB_FILE_\d+_SUM_COLUMN$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.sumColumn = value;
        } else if (field.match(/^TAB_FILE_\d+_WRITE_BACK$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.writeBack = value.toUpperCase() === 'TRUE';
        } else if (field.match(/^TAB_FILE_\d+_COMMENTS_ENABLED$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab) tab.commentsEnabled = value.toUpperCase() === 'TRUE';
        } else if (field.match(/^TAB_FILE_\d+_COLUMN_ORDER$/)) {
            const tabNum = field.split('_')[2];
            const tab = STATE.tabs.find(t => t.key === `TAB_${tabNum}`);
            if (tab && value) {
                tab.columnOrder = value.split(',').map(c => c.trim());
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
        CONFIG.primaryKey = document.getElementById('configPrimaryKey').value.trim();
        CONFIG.theme = document.getElementById('configTheme').value;
        
        const tabConfigs = document.querySelectorAll('.tab-config-item');
        STATE.tabs = [];
        
        tabConfigs.forEach((item, index) => {
            const displayName = item.querySelector('.tab-display-input').value.trim();
            const filename = item.querySelector('.tab-filename-input').value.trim();
            const sumColumn = item.querySelector('.tab-sum-input').value.trim();
            const writeBack = item.querySelector('.tab-writeback-input').checked;
            const commentsEnabled = item.querySelector('.tab-comments-input').checked;
            const columnOrder = item.querySelector('.tab-columns-input').value.trim();
            
            if (displayName && filename) {
                STATE.tabs.push({
                    key: `TAB_${index + 1}`,
                    name: filename.replace('.csv', ''),
                    displayName: displayName,
                    filename: filename,
                    sumColumn: sumColumn,
                    writeBack: writeBack,
                    commentsEnabled: commentsEnabled,
                    columnOrder: columnOrder ? columnOrder.split(',').map(c => c.trim()) : []
                });
            }
        });
        
        const configRows = [
            { Fields: 'WORKSHEET_NAME', Values: CONFIG.worksheetName },
            { Fields: 'CONFLUENCE_BASE_URL', Values: CONFIG.confluenceBaseUrl },
            { Fields: 'PAGEID', Values: CONFIG.pageId },
            { Fields: 'PRIMARY_KEY', Values: CONFIG.primaryKey },
            { Fields: 'THEME', Values: CONFIG.theme }
        ];
        
        STATE.tabs.forEach((tab, index) => {
            configRows.push({ Fields: `TAB_FILE_${index + 1}`, Values: tab.displayName });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_NAME`, Values: tab.filename });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_SUM_COLUMN`, Values: tab.sumColumn });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_WRITE_BACK`, Values: tab.writeBack ? 'TRUE' : 'FALSE' });
            configRows.push({ Fields: `TAB_FILE_${index + 1}_COMMENTS_ENABLED`, Values: tab.commentsEnabled ? 'TRUE' : 'FALSE' });
            if (tab.columnOrder.length > 0) {
                configRows.push({ Fields: `TAB_FILE_${index + 1}_COLUMN_ORDER`, Values: tab.columnOrder.join(',') });
            }
        });
        
        const csvContent = CSV.generate(configRows);
        await uploadAttachment(CONFIG.worksheetUIFile, csvContent);
        
        showToast('Configuration saved!', 'success');
        closeConfigModal();
        await init();
        
    } catch (error) {
        console.error('Save config error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
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
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
            </svg>
            ${tab.displayName}
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
        STATE.sortColumn = null;      // FIX 3: Reset sort on tab switch
        STATE.sortDirection = 'asc';  // FIX 3: Reset sort on tab switch
        
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
    
    let comments = [];
    let commentColumns = [];
    
    if (tab.commentsEnabled) {
        try {
            const commentText = await fetchAttachment(CONFIG.commentDBFile);
            comments = CSV.parse(commentText);
            
            if (comments.length > 0) {
                commentColumns = Object.keys(comments[0]).filter(col => col !== 'PRIMARY_KEY');
            } else {
                // FIX 7: Initialize with default columns for empty commentdb
                commentColumns = ['USERNAME', 'UPDATED TIME'];
            }
        } catch (error) {
            console.warn('No comments file');
            // FIX 7: Allow comments even when file doesn't exist
            commentColumns = ['USERNAME', 'UPDATED TIME'];
            comments = [];
        }
    } else {
        commentColumns = [];
    }
    
    // FIX 4: Store comment columns per tab instead of globally
    STATE.allData[tabName] = {
        rawData: data,
        filteredData: [...data],
        comments: comments,
        commentColumns: commentColumns  // Store per tab
    };
    
    // Keep global for backward compatibility but use tab-specific in new code
    STATE.commentColumns = commentColumns;
}

// ===== TABLE RENDERING (with column order fix) =====
function renderTable() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const data = tabData.filteredData;
    const comments = tabData.comments;
    
    // Get all data columns
    const allDataColumns = data.length > 0 ? Object.keys(data[0]) : [];
    
    // Build column order: specified first, then remaining
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        // Add specified columns first
        orderedColumns = [...tab.columnOrder];
        // Add remaining columns that weren't specified
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    // Prepend comment columns if enabled
    // FIX 4: Use tab-specific comment columns
    const commentColumns = tabData.commentColumns || [];
    const allColumns = tab.commentsEnabled ? [...commentColumns, ...orderedColumns] : orderedColumns;
    
    const thead = document.getElementById('tableHeader');
    thead.innerHTML = `
        <tr>
            <th><input type="checkbox" id="headerCheckbox" class="row-checkbox"></th>
            ${allColumns.map(col => {
                // FIX 3: Add sortable headers (not for comment columns)
                const isSortable = !commentColumns.includes(col);
                const sortIcon = isSortable ? `
                    <span class="sort-icon ${STATE.sortColumn === col ? 'active' : ''}" 
                          onclick="sortTable('${col}')" title="Sort by ${col}">
                        ${STATE.sortColumn === col ? 
                            (STATE.sortDirection === 'asc' ? '▲' : '▼') : '⇅'}
                    </span>
                ` : '';
                return `<th>${col} ${sortIcon}</th>`;
            }).join('')}
        </tr>
    `;
    
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
    
    pageData.forEach((row, index) => {
        const globalIndex = startIndex + index;
        const primaryKeyValue = row[CONFIG.primaryKey];
        const isSelected = STATE.selectedRows.has(globalIndex);
        
        const tr = document.createElement('tr');
        tr.className = isSelected ? 'selected' : '';
        tr.dataset.index = globalIndex;
        
        const checkboxTd = document.createElement('td');
        checkboxTd.innerHTML = `<input type="checkbox" class="row-checkbox" ${isSelected ? 'checked' : ''}>`;
        checkboxTd.querySelector('.row-checkbox').addEventListener('change', function() {
            toggleRowSelection(globalIndex, this.checked);
        });
        tr.appendChild(checkboxTd);
        
        if (tab.commentsEnabled) {
            // FIX 4: Use tab-specific comment columns
            const commentColumns = tabData.commentColumns || [];
            commentColumns.forEach(col => {
                const td = document.createElement('td');
                const comment = comments.find(c => c.PRIMARY_KEY == primaryKeyValue);
                td.textContent = comment ? (comment[col] || '') : '';
                tr.appendChild(td);
            });
        }
        
        orderedColumns.forEach(col => {
            const td = document.createElement('td');
            const value = row[col];
            
            if (col === tab.sumColumn && !isNaN(value)) {
                td.textContent = formatNumber(value);
            } else {
                td.textContent = value ?? '';
            }
            
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });
    
    renderPagination(totalPages, startIndex, endIndex, totalRows);
}

// FIX 3: Add sorting function
function sortTable(column) {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    // Toggle sort direction if same column, otherwise reset to asc
    if (STATE.sortColumn === column) {
        STATE.sortDirection = STATE.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        STATE.sortColumn = column;
        STATE.sortDirection = 'asc';
    }
    
    // Sort the filtered data
    tabData.filteredData.sort((a, b) => {
        let valA = a[column];
        let valB = b[column];
        
        // Try to parse as numbers if possible
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        
        if (!isNaN(numA) && !isNaN(numB)) {
            valA = numA;
            valB = numB;
        } else {
            // String comparison
            valA = String(valA || '').toLowerCase();
            valB = String(valB || '').toLowerCase();
        }
        
        if (valA < valB) return STATE.sortDirection === 'asc' ? -1 : 1;
        if (valA > valB) return STATE.sortDirection === 'asc' ? 1 : -1;
        return 0;
    });
    
    // Clear selection and reset page when sorting
    STATE.selectedRows.clear();
    STATE.currentPage = 1;
    
    // Re-render table with sorted data
    renderTable();
    updateSummary();
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


// ===== FILTERING (with CSS.escape fix) =====
function applyFilters() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    let filteredData = [...tabData.rawData];
    
    Object.keys(STATE.activeFilters).forEach(column => {
        const filterValues = STATE.activeFilters[column];
        if (filterValues && filterValues.length > 0) {
            filteredData = filteredData.filter(row => {
                const value = row[column]?.toString() || '';
                return filterValues.includes(value);
            });
        }
    });
    
    if (STATE.searchQuery) {
        const query = STATE.searchQuery.toLowerCase();
        filteredData = filteredData.filter(row => {
            return Object.values(row).some(value => {
                return value?.toString().toLowerCase().includes(query);
            });
        });
    }
    
    tabData.filteredData = filteredData;
    STATE.currentPage = 1;
    
    STATE.tabFilters[STATE.currentTab] = STATE.activeFilters;
    STATE.tabSearch[STATE.currentTab] = STATE.searchQuery;
}

function openFilterModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.rawData;
    if (data.length === 0) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    
    // Get all columns
    const allDataColumns = Object.keys(data[0]);
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        orderedColumns = [...tab.columnOrder];
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    const container = document.getElementById('filterColumnsContainer');
    container.innerHTML = '';
    
    orderedColumns.forEach(column => {
        const uniqueValues = [...new Set(data.map(row => row[column]?.toString() || ''))].sort();
        const isActive = STATE.activeFilters[column] && STATE.activeFilters[column].length > 0;
        
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-column-group' + (isActive ? ' active' : '');
        
        // Use CSS.escape for special characters in column names
        const escapedColumn = CSS.escape(column);
        
        filterGroup.innerHTML = `
            <label>${column}</label>
            <select multiple class="filter-select" data-column="${column}">
                ${uniqueValues.map(val => {
                    const isSelected = STATE.activeFilters[column]?.includes(val);
                    return `<option value="${val}" ${isSelected ? 'selected' : ''}>${val}</option>`;
                }).join('')}
            </select>
            <input type="text" class="filter-input" placeholder="Or type values (comma-separated)..." 
                   data-column="${column}" value="${STATE.activeFilters[column]?.join(', ') || ''}">
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
        
        // FIX: Use attribute selector with proper escaping
        const escapedColumn = CSS.escape(column);
        const textInput = document.querySelector(`.filter-input[data-column="${escapedColumn}"]`);
        
        if (textInput && textInput.value.trim()) {
            const manualValues = textInput.value.split(',').map(v => v.trim()).filter(v => v);
            selectedOptions.push(...manualValues);
        }
        
        if (selectedOptions.length > 0) {
            STATE.activeFilters[column] = [...new Set(selectedOptions)];
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

// ===== SELECTION =====
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

function updateButtonVisibility() {
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    if (!tab) return;
    
    const selectedCount = STATE.selectedRows.size;
    
    // Write-back buttons
    const editBtn = document.getElementById('editRowButton');
    const addBtn = document.getElementById('addRowButton');
    const deleteBtn = document.getElementById('deleteRowsButton');
    
    if (tab.writeBack) {
        editBtn.style.display = 'inline-flex';
        addBtn.style.display = 'inline-flex';
        deleteBtn.style.display = 'inline-flex';
        
        editBtn.disabled = selectedCount !== 1;
        deleteBtn.disabled = selectedCount === 0;
    } else {
        editBtn.style.display = 'none';
        addBtn.style.display = 'none';
        deleteBtn.style.display = 'none';
    }
    
    // Comments button
    const commentsBtn = document.getElementById('addCommentsButton');
    if (tab.commentsEnabled) {
        commentsBtn.style.display = 'inline-flex';
        commentsBtn.disabled = selectedCount === 0;
    } else {
        commentsBtn.style.display = 'none';
    }
}

// ===== SUMMARY =====
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
        filteredData.forEach((row, index) => {
            if (STATE.selectedRows.has(index)) {
                const value = parseFloat(row[tab.sumColumn]);
                if (!isNaN(value)) {
                    sum += value;
                }
            }
        });
    }
    
    // FIX 6: Make negative sums red
    const sumElement = document.getElementById('sumValue');
    sumElement.textContent = formatNumber(sum);
    sumElement.style.color = sum < 0 ? 'var(--danger)' : '';  // Red for negative
}

// ===== COMMENTS (with dynamic column save fix) =====
function openCommentsModal() {
    const selectedCount = STATE.selectedRows.size;
    if (selectedCount === 0) {
        showToast('Please select at least one row', 'warning');
        return;
    }
    
    document.getElementById('commentRowCount').textContent = selectedCount;
    
    const container = document.getElementById('commentFormContainer');
    container.innerHTML = '';
    
    // FIX 4 & 7: Use tab-specific comment columns
    const tabData = STATE.allData[STATE.currentTab];
    const commentColumns = tabData.commentColumns || ['USERNAME', 'UPDATED TIME'];
    
    // FIX 8: Create fields for editable columns only
    const editableColumns = commentColumns.filter(col => 
        col !== 'PRIMARY_KEY' && col !== 'USERNAME' && col !== 'UPDATED TIME'
    );
    
    if (editableColumns.length === 0) {
        // FIX 7: If no custom columns yet, allow adding a default one
        container.innerHTML = `
            <div class="form-group">
                <label>Comment</label>
                <textarea class="form-control comment-field" data-column="COMMENT" rows="3"></textarea>
            </div>
        `;
    } else {
        editableColumns.forEach(col => {
            const formGroup = document.createElement('div');
            formGroup.className = 'form-group';
            formGroup.innerHTML = `
                <label>${col}</label>
                <textarea class="form-control comment-field" data-column="${col}" rows="3"></textarea>
            `;
            container.appendChild(formGroup);
        });
    }
    
    // FIX 8: Show non-editable fields info
    const infoDiv = document.createElement('div');
    infoDiv.className = 'comment-info';
    infoDiv.style.marginTop = '10px';
    infoDiv.style.fontSize = '12px';
    infoDiv.style.color = 'var(--text-muted)';
    infoDiv.innerHTML = `
        <strong>Auto-filled:</strong> PRIMARY_KEY (${CONFIG.primaryKey}), 
        USERNAME (${STATE.currentUser}), 
        UPDATED TIME (Current time)
    `;
    container.appendChild(infoDiv);
    
    document.getElementById('commentModal').classList.add('show');
}

function closeCommentModal() {
    document.getElementById('commentModal').classList.remove('show');
}

async function saveComments() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        if (!tabData) return;
        
        // FIX: Dynamically collect ALL comment fields (including newly added columns)
        const commentData = {};
        document.querySelectorAll('#commentFormContainer .comment-field').forEach(field => {
            const column = field.dataset.column;
            commentData[column] = field.value;
        });
        
        commentData['USERNAME'] = STATE.currentUser;
        commentData['UPDATED TIME'] = new Date().toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
        
        const selectedData = tabData.filteredData.filter((row, index) => STATE.selectedRows.has(index));
        
        selectedData.forEach(row => {
            const primaryKeyValue = row[CONFIG.primaryKey];
            const commentIndex = tabData.comments.findIndex(c => c.PRIMARY_KEY == primaryKeyValue);
            
            const fullComment = {
                PRIMARY_KEY: primaryKeyValue,
                ...commentData
            };
            
            if (commentIndex >= 0) {
                tabData.comments[commentIndex] = fullComment;
            } else {
                tabData.comments.push(fullComment);
            }
        });
        
        // FIX 7: Initialize comments if empty and update columns if new ones added
        if (!tabData.comments || tabData.comments.length === 0) {
            tabData.comments = [];
        }
        
        // Update columns if new ones were added (like COMMENT)
        const newColumns = Object.keys(commentData);
        newColumns.forEach(col => {
            if (col !== 'PRIMARY_KEY' && !tabData.commentColumns.includes(col)) {
                tabData.commentColumns.push(col);
            }
        });
        
        const csvContent = CSV.generate(tabData.comments);
        await uploadAttachment(CONFIG.commentDBFile, csvContent);
        
        showToast('Comments saved!', 'success');
        closeCommentModal();
        
        // Reload tab to show new comment columns if any
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Save comments error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}


// ===== EDIT ROW =====
function openEditRowModal() {
    if (STATE.selectedRows.size !== 1) {
        showToast('Please select exactly one row', 'warning');
        return;
    }
    
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const selectedIndex = Array.from(STATE.selectedRows)[0];
    const row = tabData.filteredData[selectedIndex];
    
    // Get all columns with proper ordering
    const allDataColumns = Object.keys(row);
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        orderedColumns = [...tab.columnOrder];
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = selectedIndex;
    container.dataset.primaryKeyValue = row[CONFIG.primaryKey];  // FIX 1: Store primary key value
    
    // FIX 5: Ensure all columns are included and preserved
    orderedColumns.forEach(column => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        const value = row[column] || '';  // Preserve empty strings
        formGroup.innerHTML = `
            <label>${column}</label>
            <input type="text" class="form-control" data-column="${column}" value="${value}">
        `;
        container.appendChild(formGroup);
    });
    
    document.getElementById('editRowTitle').innerHTML = `
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
        </svg>
        Edit Row
    `;
    document.getElementById('editRowModal').classList.add('show');
}

function closeEditRowModal() {
    document.getElementById('editRowModal').classList.remove('show');
}

async function saveEditedRow() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        const primaryKeyValue = container.dataset.primaryKeyValue;  // FIX 1: Get stored primary key
        
        const newRow = {};
        
        // FIX 5: Collect all values including those that might be empty
        container.querySelectorAll('.form-control').forEach(input => {
            newRow[input.dataset.column] = input.value;
        });
        
        // FIX 1: Only update the specific row with matching primary key
        let updatedCount = 0;
        tabData.rawData = tabData.rawData.map(r => {
            if (r[CONFIG.primaryKey] === primaryKeyValue) {
                updatedCount++;
                return newRow;
            }
            return r;  // Return unchanged row
        });
        
        if (updatedCount === 0) {
            throw new Error('Row not found for update');
        } else if (updatedCount > 1) {
            console.warn(`Updated ${updatedCount} rows with same primary key`);
        }
        
        const csvContent = CSV.generate(tabData.rawData);
        await uploadAttachment(tab.filename, csvContent);
        await logAudit('EDIT', STATE.currentTab, newRow[CONFIG.primaryKey]);
        
        showToast('Row updated successfully!', 'success');
        closeEditRowModal();
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Edit error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ADD ROW =====
function openAddRowModal() {
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    if (!tabData) return;
    
    const allDataColumns = tabData.rawData.length > 0 ? Object.keys(tabData.rawData[0]) : [];
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        orderedColumns = [...tab.columnOrder];
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = '-1';
    
    orderedColumns.forEach(column => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        formGroup.innerHTML = `
            <label>${column}</label>
            <input type="text" class="form-control" data-column="${column}" value="">
        `;
        container.appendChild(formGroup);
    });
    
    document.getElementById('editRowTitle').innerHTML = `
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="12" y1="5" x2="12" y2="19"></line>
            <line x1="5" y1="12" x2="19" y2="12"></line>
        </svg>
        Add New Row
    `;
    document.getElementById('editRowModal').classList.add('show');
}

async function saveNewRow() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        
        if (rowIndex === -1) {
            const newRow = {};
            container.querySelectorAll('.form-control').forEach(input => {
                newRow[input.dataset.column] = input.value;
            });
            
            tabData.rawData.push(newRow);
            
            const csvContent = CSV.generate(tabData.rawData);
            await uploadAttachment(tab.filename, csvContent);
            await logAudit('ADD', STATE.currentTab, newRow[CONFIG.primaryKey]);
            
            showToast('Row added!', 'success');
            closeEditRowModal();
            
            delete STATE.allData[STATE.currentTab];
            await switchTab(STATE.currentTab);
        } else {
            await saveEditedRow();
        }
        
    } catch (error) {
        console.error('Add error:', error);
        showToast('Failed to add: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== DELETE ROWS (with confirm modal) =====
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
            const selectedData = tabData.filteredData.filter((row, index) => STATE.selectedRows.has(index));
            const primaryKeysToDelete = selectedData.map(row => row[CONFIG.primaryKey]);
            
            tabData.rawData = tabData.rawData.filter(r => !primaryKeysToDelete.includes(r[CONFIG.primaryKey]));
            
            const csvContent = CSV.generate(tabData.rawData);
            await uploadAttachment(tab.filename, csvContent);
            await logAudit('DELETE', STATE.currentTab, primaryKeysToDelete.join(','));
            
            showToast('Rows deleted!', 'success');
            
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
        
        try {
            const auditText = await fetchAttachment(CONFIG.auditLogFile);
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
        await uploadAttachment(CONFIG.auditLogFile, csvContent);
    } catch (error) {
        console.warn('Audit log failed:', error);
    }
}

// ===== EXPORT =====
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

// ===== CONFIG MODAL =====
function openConfigModal() {
    document.getElementById('configWorksheetName').value = CONFIG.worksheetName || 'Worksheet';
    document.getElementById('configBaseUrl').value = CONFIG.confluenceBaseUrl || '';
    document.getElementById('configPageId').value = CONFIG.pageId || '';
    document.getElementById('configPrimaryKey').value = CONFIG.primaryKey || 'ITEM_ID';
    document.getElementById('configTheme').value = CONFIG.theme || 'auto';
    
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
                <input type="text" class="form-control tab-display-input" value="${tab.displayName}" placeholder="Sales Data">
            </div>
            <div class="form-group">
                <label>CSV Filename</label>
                <input type="text" class="form-control tab-filename-input" value="${tab.filename}" placeholder="sales_data.csv">
            </div>
            <div class="form-group">
                <label>Sum Column</label>
                <input type="text" class="form-control tab-sum-input" value="${tab.sumColumn || ''}" placeholder="Amount">
            </div>
            <div class="form-group">
                <label>Column Order (comma-separated - these show first, rest follow)</label>
                <input type="text" class="form-control tab-columns-input" value="${tab.columnOrder.join(',')}" placeholder="ID,Name,Amount">
            </div>
            <div class="form-group">
                <label>
                    <input type="checkbox" class="tab-writeback-input" ${tab.writeBack ? 'checked' : ''}>
                    Enable Write-Back (Add/Edit/Delete)
                </label>
            </div>
            <div class="form-group">
                <label>
                    <input type="checkbox" class="tab-comments-input" ${tab.commentsEnabled ? 'checked' : ''}>
                    Enable Comments
                </label>
            </div>
        `;
        container.appendChild(tabConfig);
    });
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
        sumColumn: '',
        writeBack: false,
        commentsEnabled: true,
        columnOrder: []
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

// ===== UTILITIES =====
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
        const totalPages = Math.ceil(tabData.filteredData.length / STATE.pageSize);
        if (STATE.currentPage < totalPages) {
            goToPage(STATE.currentPage + 1);
        }
    });
    
    document.getElementById('pageSizeSelect').addEventListener('change', function() {
        STATE.pageSize = this.value;
        STATE.currentPage = 1;
        renderTable();
    });
    
    document.getElementById('addCommentsButton').addEventListener('click', openCommentsModal);
    document.getElementById('saveCommentButton').addEventListener('click', saveComments);
    
    document.getElementById('editRowButton').addEventListener('click', openEditRowModal);
    document.getElementById('addRowButton').addEventListener('click', openAddRowModal);
    document.getElementById('deleteRowsButton').addEventListener('click', deleteSelectedRows);
    document.getElementById('saveRowButton').addEventListener('click', saveNewRow);
    
    window.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal')) {
            e.target.classList.remove('show');
        }
    });
}

// ===== GLOBAL FUNCTIONS =====
window.closeConfigModal = closeConfigModal;
window.closeFilterModal = closeFilterModal;
window.clearFilters = clearFilters;
window.closeCommentModal = closeCommentModal;
window.closeEditRowModal = closeEditRowModal;
window.closeConfirmModal = closeConfirmModal;
window.removeTabConfig = removeTabConfig;

console.log('Worksheet Tool (Final Version) loaded!');
