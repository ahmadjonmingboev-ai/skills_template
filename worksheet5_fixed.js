// WORKSHEET TOOL - Final Version with All 8 Fixes
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

// ===== CONFLUENCE API =====
const API = {
    baseUrl: null,
    pageId: null,
    
    async init(baseUrl, pageId) {
        this.baseUrl = baseUrl;
        this.pageId = pageId;
    },
    
    async makeRequest(url, method = 'GET', body = null) {
        const options = {
            method: method,
            headers: {
                'Accept': 'application/json',
                'X-Atlassian-Token': 'nocheck'
            }
        };
        
        if (body) {
            if (body instanceof FormData) {
                options.body = body;
            } else {
                options.headers['Content-Type'] = 'application/json';
                options.body = JSON.stringify(body);
            }
        }
        
        const response = await fetch(url, options);
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        return response;
    }
};

// ===== STATE MANAGEMENT =====
const STATE = {
    tabs: [],
    currentTab: null,
    allData: {},
    selectedRows: new Set(),
    currentPage: 1,
    pageSize: '50',
    currentUser: 'Unknown User',
    sortColumn: null,        // FIX 3: Add sorting state
    sortDirection: 'asc'     // FIX 3: Add sorting state
};

const CONFIG = {
    primaryKey: null,
    auditLogFile: 'AuditLog.csv',
    commentDBFile: 'commentdb.csv'
};

// ===== INITIALIZE =====
async function initialize() {
    try {
        showLoading(true);
        STATE.currentUser = await getCurrentUser();
        await loadConfiguration();
        await loadTabsData();
        renderTabs();
        if (STATE.tabs.length > 0) {
            await switchTab(STATE.tabs[0].name);
        }
        applyTheme('light');
    } catch (error) {
        console.error('Init error:', error);
        showToast('Initialization failed: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function getCurrentUser() {
    try {
        const response = await fetch('/rest/api/user/current');
        const user = await response.json();
        return user.displayName || user.username || 'Unknown User';
    } catch {
        return 'Unknown User';
    }
}

// ===== CONFIGURATION =====
async function loadConfiguration() {
    try {
        const configText = await fetchAttachment('worksheetui.xlsx');
        const lines = configText.split('\n').filter(l => l.trim());
        
        lines.slice(1).forEach(line => {
            const [key, value] = line.split(',').map(s => s.trim());
            
            if (key === 'Page ID') {
                API.pageId = value;
            } else if (key === 'Confluence Base URL') {
                API.baseUrl = value;
            } else if (key === 'Primary Key') {
                CONFIG.primaryKey = value;
            }
        });
        
        if (!CONFIG.primaryKey) {
            CONFIG.primaryKey = 'ID';
        }
    } catch (error) {
        console.error('Config load error:', error);
        throw new Error('Failed to load configuration');
    }
}

async function loadTabsData() {
    try {
        const response = await API.makeRequest(
            `${API.baseUrl}/rest/api/content/${API.pageId}/child/attachment`,
            'GET'
        );
        const attachments = await response.json();
        
        const tabFile = attachments.results.find(a => a.title === 'worksheetui_tabs.xlsx');
        if (!tabFile) {
            throw new Error('Tabs configuration not found');
        }
        
        const tabData = await fetch(API.baseUrl + tabFile._links.download).then(r => r.text());
        const lines = tabData.split('\n').filter(l => l.trim());
        
        STATE.tabs = [];
        lines.slice(1).forEach(line => {
            const parts = line.split(',').map(s => s.trim());
            if (parts.length >= 4) {
                STATE.tabs.push({
                    name: parts[0],
                    filename: parts[1],
                    color: parts[2],
                    writeBack: parts[3]?.toLowerCase() === 'true',
                    sumColumn: parts[4] || null,
                    commentsEnabled: parts[5]?.toLowerCase() === 'true',
                    columnOrder: parts[6] ? parts[6].split('|').filter(c => c) : [],
                    filterableColumns: parts[7] ? parts[7].split('|').filter(c => c) : []
                });
            }
        });
    } catch (error) {
        console.error('Tabs load error:', error);
        STATE.tabs = [{
            name: 'Default',
            filename: 'data.csv',
            color: '#3b82f6',
            writeBack: false,
            sumColumn: null,
            commentsEnabled: false,
            columnOrder: [],
            filterableColumns: []
        }];
    }
}

// ===== ATTACHMENT OPERATIONS =====
async function fetchAttachment(filename) {
    const response = await API.makeRequest(
        `${API.baseUrl}/rest/api/content/${API.pageId}/child/attachment`,
        'GET'
    );
    const data = await response.json();
    const attachment = data.results.find(a => a.title === filename);
    
    if (!attachment) {
        throw new Error(`Attachment ${filename} not found`);
    }
    
    const content = await fetch(API.baseUrl + attachment._links.download);
    return await content.text();
}

async function uploadAttachment(filename, content) {
    const attachmentsResponse = await API.makeRequest(
        `${API.baseUrl}/rest/api/content/${API.pageId}/child/attachment`,
        'GET'
    );
    const attachments = await attachmentsResponse.json();
    const existing = attachments.results.find(a => a.title === filename);
    
    const blob = new Blob([content], { type: 'text/csv' });
    const formData = new FormData();
    formData.append('file', blob, filename);
    formData.append('minorEdit', 'true');
    
    let url = `${API.baseUrl}/rest/api/content/${API.pageId}/child/attachment`;
    if (existing) {
        url += `/${existing.id}/data`;
    }
    
    await API.makeRequest(url, 'POST', formData);
}

// ===== TAB MANAGEMENT =====
async function switchTab(tabName) {
    if (!STATE.allData[tabName]) {
        await loadTabData(tabName);
    }
    
    STATE.currentTab = tabName;
    STATE.currentPage = 1;
    STATE.selectedRows.clear();
    STATE.sortColumn = null;  // FIX 3: Reset sort on tab switch
    STATE.sortDirection = 'asc';
    
    renderTabs();
    renderControls();
    renderTable();
    updateButtonVisibility();
    updateSummary();
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
            
            // FIX 7 & 8: Handle empty commentdb but ensure proper structure
            if (comments.length > 0) {
                commentColumns = Object.keys(comments[0]).filter(col => col !== 'PRIMARY_KEY');
            } else {
                // Create default structure for empty commentdb
                commentColumns = ['USERNAME', 'UPDATED TIME'];
            }
        } catch (error) {
            console.warn('No comments file or empty file');
            // FIX 7: Allow comments even when file doesn't exist
            commentColumns = ['USERNAME', 'UPDATED TIME'];
            comments = [];
        }
    }
    
    // FIX 4: Store comment columns per tab instead of globally
    STATE.allData[tabName] = {
        rawData: data,
        filteredData: [...data],
        comments: comments,
        commentColumns: commentColumns  // Store per tab
    };
}

function renderTabs() {
    const container = document.getElementById('tabContainer');
    
    container.innerHTML = STATE.tabs.map(tab => `
        <button 
            class="tab ${tab.name === STATE.currentTab ? 'active' : ''}"
            onclick="switchTab('${tab.name}')"
            style="${tab.name === STATE.currentTab ? `border-bottom-color: ${tab.color};` : ''}"
        >
            ${tab.name}
        </button>
    `).join('');
}

// ===== CONTROLS =====
function renderControls() {
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const allColumns = Object.keys(tabData.rawData[0] || {});
    
    const filterableColumns = tab.filterableColumns.length > 0 
        ? tab.filterableColumns 
        : allColumns;
    
    document.getElementById('filterableColumnsList').innerHTML = filterableColumns.map(col => `
        <button class="filter-chip" onclick="toggleColumnFilter('${col}')">
            ${col}
        </button>
    `).join('');
}

// ===== TABLE RENDERING =====
function renderTable() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const allDataColumns = Object.keys(tabData.rawData[0] || {});
    
    // FIX 5: Preserve all columns when ordering
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        // Start with ordered columns
        orderedColumns = [...tab.columnOrder].filter(col => allDataColumns.includes(col));
        // Add remaining columns that aren't in the order
        const remainingColumns = allDataColumns.filter(col => !tab.columnOrder.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    // FIX 4: Use tab-specific comment columns
    const commentColumns = tabData.commentColumns || [];
    const allColumns = tab.commentsEnabled ? [...commentColumns, ...orderedColumns] : orderedColumns;
    
    const pageSize = parseInt(STATE.pageSize);
    const startIndex = (STATE.currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, tabData.filteredData.length);
    const pageData = tabData.filteredData.slice(startIndex, endIndex);
    
    const table = document.getElementById('dataTable');
    
    // Create header with sorting
    let headerHTML = `
        <thead>
            <tr>
                <th class="checkbox-column">
                    <input type="checkbox" class="header-checkbox" 
                           onchange="this.checked ? selectAllRows() : unselectAllRows()">
                </th>
    `;
    
    // FIX 3: Add sortable headers
    allColumns.forEach(column => {
        const isSortable = !commentColumns.includes(column);
        const sortIcon = isSortable ? `
            <span class="sort-icon ${STATE.sortColumn === column ? 'active' : ''}" 
                  onclick="sortTable('${column}')" title="Sort by ${column}">
                ${STATE.sortColumn === column ? 
                    (STATE.sortDirection === 'asc' ? '▲' : '▼') : '⇅'}
            </span>
        ` : '';
        
        headerHTML += `<th>${column} ${sortIcon}</th>`;
    });
    
    headerHTML += `</tr></thead><tbody>`;
    
    // Render rows
    pageData.forEach((row, pageIndex) => {
        const actualIndex = startIndex + pageIndex;
        const isSelected = STATE.selectedRows.has(actualIndex);
        
        headerHTML += `
            <tr data-index="${actualIndex}" class="${isSelected ? 'selected' : ''}">
                <td class="checkbox-column">
                    <input type="checkbox" class="row-checkbox" ${isSelected ? 'checked' : ''}
                           onchange="toggleRowSelection(${actualIndex}, this.checked)">
                </td>
        `;
        
        allColumns.forEach(column => {
            let value = '';
            
            if (commentColumns.includes(column)) {
                const comment = tabData.comments.find(c => c.PRIMARY_KEY == row[CONFIG.primaryKey]);
                value = comment ? (comment[column] || '') : '';
            } else {
                value = row[column] || '';
            }
            
            const isNumeric = !isNaN(parseFloat(value)) && column === tab.sumColumn;
            
            headerHTML += `<td ${isNumeric ? 'class="numeric"' : ''}>${
                isNumeric ? formatNumber(value) : value
            }</td>`;
        });
        
        headerHTML += `</tr>`;
    });
    
    headerHTML += `</tbody>`;
    table.innerHTML = headerHTML;
    
    renderPagination(tabData.filteredData.length);
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
    
    // Re-render table with sorted data
    renderTable();
    updateSummary();
}

function renderPagination(totalRows) {
    const pageSize = parseInt(STATE.pageSize);
    const totalPages = Math.ceil(totalRows / pageSize);
    
    const paginationHTML = `
        <div class="pagination-info">
            Showing ${Math.min((STATE.currentPage - 1) * pageSize + 1, totalRows)} - 
            ${Math.min(STATE.currentPage * pageSize, totalRows)} of ${totalRows}
        </div>
        <div class="pagination-controls">
            <button class="btn btn-sm" onclick="changePage(1)" 
                    ${STATE.currentPage === 1 ? 'disabled' : ''}>First</button>
            <button class="btn btn-sm" onclick="changePage(${STATE.currentPage - 1})" 
                    ${STATE.currentPage === 1 ? 'disabled' : ''}>Previous</button>
            <span class="page-numbers">
                Page ${STATE.currentPage} of ${totalPages}
            </span>
            <button class="btn btn-sm" onclick="changePage(${STATE.currentPage + 1})" 
                    ${STATE.currentPage === totalPages ? 'disabled' : ''}>Next</button>
            <button class="btn btn-sm" onclick="changePage(${totalPages})" 
                    ${STATE.currentPage === totalPages ? 'disabled' : ''}>Last</button>
        </div>
        <select class="page-size-select" value="${STATE.pageSize}" 
                onchange="changePageSize(this.value)">
            <option value="10">10 rows</option>
            <option value="25">25 rows</option>
            <option value="50">50 rows</option>
            <option value="100">100 rows</option>
        </select>
    `;
    
    document.getElementById('paginationControls').innerHTML = paginationHTML;
    document.querySelector('.page-size-select').value = STATE.pageSize;
}

function changePage(page) {
    STATE.currentPage = page;
    renderTable();
}

function changePageSize(size) {
    STATE.pageSize = size;
    STATE.currentPage = 1;
    renderTable();
}

// ===== SEARCH & FILTER =====
function handleSearch(value) {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const searchTerm = value.toLowerCase();
    
    if (searchTerm) {
        tabData.filteredData = tabData.rawData.filter(row => 
            Object.values(row).some(val => 
                String(val).toLowerCase().includes(searchTerm)
            )
        );
    } else {
        tabData.filteredData = [...tabData.rawData];
    }
    
    STATE.currentPage = 1;
    STATE.selectedRows.clear();
    renderTable();
    updateSummary();
}

function openFilterModal() {
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const tabData = STATE.allData[STATE.currentTab];
    
    const filterableColumns = tab.filterableColumns.length > 0 
        ? tab.filterableColumns 
        : Object.keys(tabData.rawData[0] || {});
    
    const container = document.getElementById('filterFormContainer');
    container.innerHTML = '';
    
    filterableColumns.forEach(column => {
        const uniqueValues = [...new Set(tabData.rawData.map(row => row[column]))].filter(v => v);
        
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        formGroup.innerHTML = `
            <label>${column}</label>
            <select class="form-control filter-select" data-column="${column}" multiple size="5">
                ${uniqueValues.map(val => `
                    <option value="${val}">${val}</option>
                `).join('')}
            </select>
            <small>Hold Ctrl/Cmd to select multiple</small>
        `;
        container.appendChild(formGroup);
    });
    
    document.getElementById('filterModal').classList.add('show');
}

function closeFilterModal() {
    document.getElementById('filterModal').classList.remove('show');
}

function applyFilters() {
    const tabData = STATE.allData[STATE.currentTab];
    const filters = {};
    
    document.querySelectorAll('.filter-select').forEach(select => {
        const column = select.dataset.column;
        const selectedOptions = Array.from(select.selectedOptions).map(opt => opt.value);
        if (selectedOptions.length > 0) {
            filters[column] = selectedOptions;
        }
    });
    
    if (Object.keys(filters).length > 0) {
        tabData.filteredData = tabData.rawData.filter(row => {
            return Object.entries(filters).every(([column, values]) => {
                return values.includes(row[column]);
            });
        });
    } else {
        tabData.filteredData = [...tabData.rawData];
    }
    
    STATE.currentPage = 1;
    STATE.selectedRows.clear();
    renderTable();
    updateSummary();
    closeFilterModal();
    showToast('Filters applied', 'success');
}

function clearFilters() {
    const tabData = STATE.allData[STATE.currentTab];
    tabData.filteredData = [...tabData.rawData];
    STATE.currentPage = 1;
    STATE.selectedRows.clear();
    renderTable();
    updateSummary();
    showToast('Filters cleared', 'info');
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
    const tabData = STATE.allData[STATE.currentTab];
    
    // FIX 7: Enable comments button even if no commentdb exists yet
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

// ===== COMMENTS (with fixes) =====
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
        // If no custom columns yet, allow adding a default one
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
    
    // Show non-editable fields info
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
        
        // FIX 8: Collect comment data with proper structure
        const commentData = {};
        document.querySelectorAll('#commentFormContainer .comment-field').forEach(field => {
            const column = field.dataset.column;
            commentData[column] = field.value;
        });
        
        // FIX 8: Add system fields
        commentData['USERNAME'] = STATE.currentUser;
        commentData['UPDATED TIME'] = new Date().toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
        
        const selectedData = tabData.filteredData.filter((row, index) => STATE.selectedRows.has(index));
        
        // FIX 7: Initialize comments if empty
        if (!tabData.comments || tabData.comments.length === 0) {
            tabData.comments = [];
            // Update columns if new ones were added
            const newColumns = Object.keys(commentData);
            if (!tabData.commentColumns.includes('COMMENT') && commentData['COMMENT']) {
                tabData.commentColumns.push('COMMENT');
            }
        }
        
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
        
        const csvContent = CSV.generate(tabData.comments);
        await uploadAttachment(CONFIG.commentDBFile, csvContent);
        
        showToast('Comments saved!', 'success');
        closeCommentModal();
        
        // Reload to show new comment columns
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Save comments error:', error);
        showToast('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}


// ===== EDIT ROW (with fix) =====
function openEditRowModal() {
    if (STATE.selectedRows.size !== 1) {
        showToast('Please select exactly one row', 'warning');
        return;
    }
    
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const selectedIndex = Array.from(STATE.selectedRows)[0];
    const row = tabData.filteredData[selectedIndex];
    
    // FIX 5: Ensure all columns are included
    const allDataColumns = Object.keys(row);
    let orderedColumns = [];
    if (tab.columnOrder && tab.columnOrder.length > 0) {
        // Include ordered columns that exist
        orderedColumns = tab.columnOrder.filter(col => allDataColumns.includes(col));
        // Add remaining columns
        const remainingColumns = allDataColumns.filter(col => !orderedColumns.includes(col));
        orderedColumns = [...orderedColumns, ...remainingColumns];
    } else {
        orderedColumns = allDataColumns;
    }
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = selectedIndex;
    container.dataset.primaryKeyValue = row[CONFIG.primaryKey];  // Store primary key
    
    orderedColumns.forEach(column => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        const value = row[column] || '';
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
        
        // Reload the tab to refresh data
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
    container.removeAttribute('data-rowIndex');
    
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
        
        const newRow = {};
        container.querySelectorAll('.form-control').forEach(input => {
            newRow[input.dataset.column] = input.value;
        });
        
        if (!newRow[CONFIG.primaryKey]) {
            throw new Error(`Primary key (${CONFIG.primaryKey}) is required`);
        }
        
        tabData.rawData.push(newRow);
        
        const csvContent = CSV.generate(tabData.rawData);
        await uploadAttachment(tab.filename, csvContent);
        await logAudit('ADD', STATE.currentTab, newRow[CONFIG.primaryKey]);
        
        showToast('New row added!', 'success');
        closeEditRowModal();
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Add error:', error);
        showToast('Failed to add: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== DELETE ROWS =====
function deleteSelectedRows() {
    const selectedCount = STATE.selectedRows.size;
    if (selectedCount === 0) {
        showToast('Please select rows to delete', 'warning');
        return;
    }
    
    document.getElementById('confirmMessage').textContent = 
        `Are you sure you want to delete ${selectedCount} row(s)? This cannot be undone.`;
    
    const confirmBtn = document.getElementById('confirmButton');
    confirmBtn.onclick = confirmDelete;
    
    document.getElementById('confirmModal').classList.add('show');
}

function closeConfirmModal() {
    document.getElementById('confirmModal').classList.remove('show');
}

async function confirmDelete() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        
        const selectedRows = tabData.filteredData.filter((row, index) => 
            STATE.selectedRows.has(index)
        );
        
        const primaryKeys = selectedRows.map(row => row[CONFIG.primaryKey]);
        
        tabData.rawData = tabData.rawData.filter(row => 
            !primaryKeys.includes(row[CONFIG.primaryKey])
        );
        
        const csvContent = CSV.generate(tabData.rawData);
        await uploadAttachment(tab.filename, csvContent);
        
        for (const pk of primaryKeys) {
            await logAudit('DELETE', STATE.currentTab, pk);
        }
        
        showToast(`${primaryKeys.length} row(s) deleted!`, 'success');
        closeConfirmModal();
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Delete error:', error);
        showToast('Failed to delete: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== EXPORT =====
async function exportToExcel() {
    try {
        const tabData = STATE.allData[STATE.currentTab];
        if (!tabData || tabData.filteredData.length === 0) {
            showToast('No data to export', 'warning');
            return;
        }
        
        const csvContent = CSV.generate(tabData.filteredData);
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${STATE.currentTab}_export_${new Date().toISOString().split('T')[0]}.csv`;
        link.click();
        
        showToast('Data exported!', 'success');
    } catch (error) {
        console.error('Export error:', error);
        showToast('Export failed', 'error');
    }
}

// ===== AUDIT LOG =====
async function logAudit(action, tabName, recordId) {
    try {
        let auditLog = [];
        
        try {
            const auditText = await fetchAttachment(CONFIG.auditLogFile);
            auditLog = CSV.parse(auditText);
        } catch (e) {
            console.log('Creating new audit log');
        }
        
        auditLog.push({
            TIMESTAMP: new Date().toISOString(),
            USER: STATE.currentUser,
            ACTION: action,
            TAB: tabName,
            RECORD_ID: recordId
        });
        
        const csvContent = CSV.generate(auditLog);
        await uploadAttachment(CONFIG.auditLogFile, csvContent);
        
    } catch (error) {
        console.error('Audit log error:', error);
    }
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

// ===== INITIALIZE ON LOAD =====
document.addEventListener('DOMContentLoaded', initialize);
