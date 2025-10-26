// ============================================
// WORKSHEET TOOL - JavaScript
// ============================================

// ===== CONFIGURATION =====
const CONFIG = {
    confluenceBaseUrl: window.location.origin,
    pageId: '',
    worksheetUIFile: 'worksheetui.xlsx',
    commentDBFile: 'commentdb.xlsx',
    auditLogFile: 'audit_log.xlsx',
    primaryKey: 'ITEM_ID',
    sumColumn: 'Amount',
    theme: 'auto',
    writeBackEnabled: false
};

// ===== STATE MANAGEMENT =====
const STATE = {
    currentTab: null,
    tabs: [],
    allData: {},  // {tabName: {rawData: [], filteredData: [], comments: []}}
    currentPage: 1,
    pageSize: 50,
    selectedRows: new Set(),
    activeFilters: {},
    searchQuery: '',
    currentUser: 'Loading...',
    columnOrders: {}  // {tabName: [col1, col2, ...]}
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', async function() {
    console.log('Worksheet Tool initializing...');
    await init();
});

async function init() {
    try {
        showLoading(true);
        
        // Get current user
        await getCurrentUser();
        
        // Load configuration
        await loadConfiguration();
        
        // Apply theme
        applyTheme(CONFIG.theme);
        
        // Setup event listeners
        setupEventListeners();
        
        // Load first tab if available
        if (STATE.tabs.length > 0) {
            await switchTab(STATE.tabs[0].name);
        } else {
            showMessage('No tabs configured. Please configure worksheet settings.', 'warning');
        }
        
        showLoading(false);
    } catch (error) {
        console.error('Initialization error:', error);
        showMessage('Failed to initialize worksheet: ' + error.message, 'error');
        showLoading(false);
    }
}

// ===== CONFLUENCE API =====
async function getCurrentUser() {
    try {
        // Try to get current user from Confluence
        const response = await fetch(`${CONFIG.confluenceBaseUrl}/rest/api/user/current`, {
            credentials: 'include'
        });
        
        if (response.ok) {
            const user = await response.json();
            STATE.currentUser = user.displayName || user.username || 'Unknown User';
        } else {
            STATE.currentUser = 'Guest User';
        }
    } catch (error) {
        console.warn('Could not get current user:', error);
        STATE.currentUser = 'Guest User';
    }
    
    document.getElementById('currentUser').textContent = STATE.currentUser;
}

async function fetchAttachment(filename) {
    try {
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
        
        // Download the file
        const downloadUrl = CONFIG.confluenceBaseUrl + attachment._links.download;
        const fileResponse = await fetch(downloadUrl, { credentials: 'include' });
        
        if (!fileResponse.ok) {
            throw new Error(`Failed to download ${filename}`);
        }
        
        const arrayBuffer = await fileResponse.arrayBuffer();
        return arrayBuffer;
    } catch (error) {
        console.error(`Error fetching ${filename}:`, error);
        throw error;
    }
}

async function uploadAttachment(filename, fileContent) {
    try {
        const formData = new FormData();
        const blob = new Blob([fileContent], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        formData.append('file', blob, filename);
        formData.append('comment', `Updated by ${STATE.currentUser} via Worksheet Tool`);
        
        const url = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
        const response = await fetch(url, {
            method: 'POST',
            credentials: 'include',
            headers: {
                'X-Atlassian-Token': 'no-check'
            },
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`Failed to upload ${filename}: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error(`Error uploading ${filename}:`, error);
        throw error;
    }
}

// ===== CONFIGURATION MANAGEMENT =====
async function loadConfiguration() {
    try {
        const configBuffer = await fetchAttachment(CONFIG.worksheetUIFile);
        const workbook = XLSX.read(configBuffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const configData = XLSX.utils.sheet_to_json(sheet, { header: ['Fields', 'Values'] });
        
        // Parse configuration
        configData.forEach(row => {
            const field = row.Fields;
            const value = row.Values;
            
            if (!field || !value) return;
            
            if (field === 'PAGEID') {
                CONFIG.pageId = value.toString();
            } else if (field === 'PRIMARY_KEY') {
                CONFIG.primaryKey = value;
            } else if (field === 'SUM_COLUMN') {
                CONFIG.sumColumn = value;
            } else if (field === 'THEME') {
                CONFIG.theme = value;
            } else if (field === 'WRITE_BACK') {
                CONFIG.writeBackEnabled = value.toString().toUpperCase() === 'TRUE';
            } else if (field.startsWith('TAB_FILE_')) {
                const tabNum = field.split('_')[2];
                const tabKey = `TAB_${tabNum}`;
                
                if (!STATE.tabs.find(t => t.key === tabKey)) {
                    STATE.tabs.push({
                        key: tabKey,
                        name: value.replace('.xlsx', ''),
                        filename: value,
                        displayName: value.replace('.xlsx', '').replace(/_/g, ' ')
                    });
                }
            } else if (field.startsWith('TAB_FILE_') && field.endsWith('_NAME')) {
                const tabNum = field.split('_')[2];
                const tabKey = `TAB_${tabNum}`;
                const tab = STATE.tabs.find(t => t.key === tabKey);
                if (tab) tab.displayName = value;
            } else if (field.startsWith('TAB_FILE_') && field.endsWith('_COLUMN_ORDER')) {
                const tabNum = field.split('_')[2];
                const tabKey = `TAB_${tabNum}`;
                const tab = STATE.tabs.find(t => t.key === tabKey);
                if (tab && value) {
                    STATE.columnOrders[tab.name] = value.split(',').map(c => c.trim());
                }
            }
        });
        
        // Render tabs
        renderTabs();
        
        // Show/hide write-back buttons
        if (CONFIG.writeBackEnabled) {
            document.getElementById('addRowButton').style.display = 'inline-flex';
        }
        
    } catch (error) {
        console.warn('Could not load configuration, using defaults:', error);
        // Show config modal to let user set up
        setTimeout(() => openConfigModal(), 500);
    }
}

async function saveConfiguration() {
    try {
        showLoading(true);
        
        // Gather config data from form
        CONFIG.pageId = document.getElementById('configPageId').value;
        CONFIG.primaryKey = document.getElementById('configPrimaryKey').value;
        CONFIG.sumColumn = document.getElementById('configSumColumn').value;
        CONFIG.theme = document.getElementById('configTheme').value;
        CONFIG.writeBackEnabled = document.getElementById('configWriteBack').checked;
        
        // Gather tab configurations
        const tabInputs = document.querySelectorAll('.tab-filename-input');
        STATE.tabs = [];
        
        tabInputs.forEach((input, index) => {
            const filename = input.value.trim();
            const nameInput = input.closest('.tab-config-item').querySelector('.tab-name-input');
            const displayName = nameInput ? nameInput.value.trim() : filename.replace('.xlsx', '');
            
            if (filename) {
                STATE.tabs.push({
                    key: `TAB_${index + 1}`,
                    name: filename.replace('.xlsx', ''),
                    filename: filename,
                    displayName: displayName
                });
            }
        });
        
        // Create Excel file with configuration
        const configRows = [
            ['Fields', 'Values'],
            ['PAGEID', CONFIG.pageId],
            ['PRIMARY_KEY', CONFIG.primaryKey],
            ['SUM_COLUMN', CONFIG.sumColumn],
            ['THEME', CONFIG.theme],
            ['WRITE_BACK', CONFIG.writeBackEnabled ? 'TRUE' : 'FALSE'],
            ['', '']
        ];
        
        STATE.tabs.forEach((tab, index) => {
            configRows.push([`TAB_FILE_${index + 1}`, tab.filename]);
            configRows.push([`TAB_FILE_${index + 1}_NAME`, tab.displayName]);
            
            if (STATE.columnOrders[tab.name]) {
                configRows.push([`TAB_FILE_${index + 1}_COLUMN_ORDER`, STATE.columnOrders[tab.name].join(',')]);
            }
        });
        
        const ws = XLSX.utils.aoa_to_sheet(configRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        await uploadAttachment(CONFIG.worksheetUIFile, wbout);
        
        showMessage('Configuration saved successfully!', 'success');
        closeConfigModal();
        
        // Re-initialize
        await init();
        
    } catch (error) {
        console.error('Error saving configuration:', error);
        showMessage('Failed to save configuration: ' + error.message, 'error');
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
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
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
        
        // Update active tab button
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tabName === tabName);
        });
        
        // Load tab data if not already loaded
        if (!STATE.allData[tabName]) {
            await loadTabData(tabName);
        }
        
        // Apply filters and render
        applyFilters();
        renderTable();
        updateSummary();
        
        showLoading(false);
    } catch (error) {
        console.error('Error switching tab:', error);
        showMessage('Failed to load tab data: ' + error.message, 'error');
        showLoading(false);
    }
}

async function loadTabData(tabName) {
    const tab = STATE.tabs.find(t => t.name === tabName);
    if (!tab) return;
    
    try {
        // Fetch Excel file
        const buffer = await fetchAttachment(tab.filename);
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet);
        
        // Load comments if exists
        let comments = [];
        try {
            const commentBuffer = await fetchAttachment(CONFIG.commentDBFile);
            const commentWorkbook = XLSX.read(commentBuffer, { type: 'array' });
            const commentSheet = commentWorkbook.Sheets[commentWorkbook.SheetNames[0]];
            comments = XLSX.utils.sheet_to_json(commentSheet);
        } catch (error) {
            console.warn('No comments file found, continuing without comments');
        }
        
        // Store data
        STATE.allData[tabName] = {
            rawData: data,
            filteredData: [...data],
            comments: comments
        };
        
    } catch (error) {
        console.error('Error loading tab data:', error);
        throw error;
    }
}


// ===== TABLE RENDERING =====
function renderTable() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.filteredData;
    const comments = tabData.comments;
    
    // Get columns to display
    let columns = STATE.columnOrders[STATE.currentTab];
    if (!columns || columns.length === 0) {
        // Use all columns from data
        columns = data.length > 0 ? Object.keys(data[0]) : [];
    }
    
    // Render header
    const thead = document.getElementById('tableHeader');
    thead.innerHTML = `
        <tr>
            <th><input type="checkbox" id="headerCheckbox" class="row-checkbox"></th>
            ${columns.map(col => `<th>${col}</th>`).join('')}
            <th>Actions</th>
        </tr>
    `;
    
    // Setup header checkbox
    document.getElementById('headerCheckbox').addEventListener('change', function() {
        if (this.checked) {
            selectAllRows();
        } else {
            unselectAllRows();
        }
    });
    
    // Calculate pagination
    const totalRows = data.length;
    const pageSize = STATE.pageSize === 'all' ? totalRows : parseInt(STATE.pageSize);
    const totalPages = Math.ceil(totalRows / pageSize);
    const startIndex = (STATE.currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalRows);
    const pageData = data.slice(startIndex, endIndex);
    
    // Render body
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (pageData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="${columns.length + 2}" style="text-align: center; padding: 40px;">
                    No data available
                </td>
            </tr>
        `;
        return;
    }
    
    pageData.forEach((row, index) => {
        const globalIndex = startIndex + index;
        const primaryKeyValue = row[CONFIG.primaryKey];
        const isSelected = STATE.selectedRows.has(globalIndex);
        const hasComment = comments.some(c => c.PRIMARY_KEY == primaryKeyValue);
        
        const tr = document.createElement('tr');
        tr.className = isSelected ? 'selected' : '';
        tr.dataset.index = globalIndex;
        
        // Checkbox column
        const checkboxTd = document.createElement('td');
        checkboxTd.innerHTML = `<input type="checkbox" class="row-checkbox" ${isSelected ? 'checked' : ''}>`;
        checkboxTd.querySelector('.row-checkbox').addEventListener('change', function() {
            toggleRowSelection(globalIndex, this.checked);
        });
        tr.appendChild(checkboxTd);
        
        // Data columns
        columns.forEach(col => {
            const td = document.createElement('td');
            const value = row[col];
            
            // Format based on column type
            if (col === CONFIG.sumColumn && !isNaN(value)) {
                td.textContent = formatCurrency(value);
            } else {
                td.textContent = value ?? '';
            }
            
            tr.appendChild(td);
        });
        
        // Actions column
        const actionsTd = document.createElement('td');
        actionsTd.innerHTML = `
            <button class="comment-btn ${hasComment ? 'has-comment' : ''}" onclick="openCommentModal(${globalIndex})">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
                </svg>
                Comment
            </button>
            ${CONFIG.writeBackEnabled ? `
                <button class="edit-btn" onclick="openEditRowModal(${globalIndex})">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                    </svg>
                </button>
            ` : ''}
        `;
        tr.appendChild(actionsTd);
        
        tbody.appendChild(tr);
    });
    
    // Render pagination
    renderPagination(totalPages, startIndex, endIndex, totalRows);
}

function renderPagination(totalPages, startIndex, endIndex, totalRows) {
    document.getElementById('pageStart').textContent = startIndex + 1;
    document.getElementById('pageEnd').textContent = endIndex;
    document.getElementById('totalRowsPage').textContent = totalRows;
    
    // Page numbers
    const pageNumbersContainer = document.getElementById('pageNumbers');
    pageNumbersContainer.innerHTML = '';
    
    const maxButtons = 7;
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
    
    // Enable/disable navigation buttons
    document.getElementById('firstPageBtn').disabled = STATE.currentPage === 1;
    document.getElementById('prevPageBtn').disabled = STATE.currentPage === 1;
    document.getElementById('nextPageBtn').disabled = STATE.currentPage === totalPages;
    document.getElementById('lastPageBtn').disabled = STATE.currentPage === totalPages;
}

function goToPage(page) {
    STATE.currentPage = page;
    renderTable();
}

// ===== FILTERING & SEARCH =====
function applyFilters() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    let filteredData = [...tabData.rawData];
    
    // Apply column filters
    Object.keys(STATE.activeFilters).forEach(column => {
        const filterValues = STATE.activeFilters[column];
        if (filterValues && filterValues.length > 0) {
            filteredData = filteredData.filter(row => {
                const value = row[column]?.toString() || '';
                return filterValues.includes(value);
            });
        }
    });
    
    // Apply search
    if (STATE.searchQuery) {
        const query = STATE.searchQuery.toLowerCase();
        filteredData = filteredData.filter(row => {
            return Object.values(row).some(value => {
                return value?.toString().toLowerCase().includes(query);
            });
        });
    }
    
    tabData.filteredData = filteredData;
    STATE.currentPage = 1; // Reset to first page
}

function openFilterModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.rawData;
    if (data.length === 0) return;
    
    const columns = STATE.columnOrders[STATE.currentTab] || Object.keys(data[0]);
    const container = document.getElementById('filterColumnsContainer');
    container.innerHTML = '';
    
    columns.forEach(column => {
        // Get unique values for this column
        const uniqueValues = [...new Set(data.map(row => row[column]?.toString() || ''))].sort();
        
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-column-group';
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
        
        // Also check text input
        const textInput = document.querySelector(`.filter-input[data-column="${column}"]`);
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
    document.getElementById('filterBadge').style.display = 'none';
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

// ===== SELECTION MANAGEMENT =====
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
}

function selectAllRows() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.filteredData;
    data.forEach((row, index) => {
        STATE.selectedRows.add(index);
    });
    
    document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = true);
    document.querySelectorAll('tbody tr').forEach(tr => tr.classList.add('selected'));
    
    updateSummary();
}

function unselectAllRows() {
    STATE.selectedRows.clear();
    document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = false);
    document.querySelectorAll('tbody tr').forEach(tr => tr.classList.remove('selected'));
    updateSummary();
}

// ===== SUMMARY & STATISTICS =====
function updateSummary() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const filteredData = tabData.filteredData;
    const totalRows = filteredData.length;
    const selectedCount = STATE.selectedRows.size;
    
    document.getElementById('totalRows').textContent = totalRows;
    document.getElementById('selectedRows').textContent = selectedCount;
    
    // Calculate sum of selected rows
    let sum = 0;
    if (CONFIG.sumColumn) {
        filteredData.forEach((row, index) => {
            if (STATE.selectedRows.has(index)) {
                const value = parseFloat(row[CONFIG.sumColumn]);
                if (!isNaN(value)) {
                    sum += value;
                }
            }
        });
    }
    
    document.getElementById('sumValue').textContent = formatCurrency(sum);
}

// ===== COMMENTS =====
function openCommentModal(rowIndex) {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const row = tabData.filteredData[rowIndex];
    const primaryKeyValue = row[CONFIG.primaryKey];
    
    // Find existing comment
    const existingComment = tabData.comments.find(c => c.PRIMARY_KEY == primaryKeyValue);
    
    document.getElementById('commentRowKey').value = primaryKeyValue;
    
    if (existingComment) {
        document.getElementById('commentRootCause').value = existingComment['ROOT CAUSE'] || '';
        document.getElementById('commentAction').value = existingComment['ACTION'] || '';
        document.getElementById('commentETA').value = existingComment['ETA'] || '';
        
        document.getElementById('commentUsername').textContent = existingComment['USERNAME'] || '';
        document.getElementById('commentTimestamp').textContent = existingComment['UPDATED TIME'] || '';
        document.getElementById('commentMeta').style.display = 'block';
    } else {
        document.getElementById('commentRootCause').value = '';
        document.getElementById('commentAction').value = '';
        document.getElementById('commentETA').value = '';
        document.getElementById('commentMeta').style.display = 'none';
    }
    
    document.getElementById('commentModal').classList.add('show');
}

function closeCommentModal() {
    document.getElementById('commentModal').classList.remove('show');
}

async function saveComment() {
    try {
        showLoading(true);
        
        const primaryKeyValue = document.getElementById('commentRowKey').value;
        const rootCause = document.getElementById('commentRootCause').value;
        const action = document.getElementById('commentAction').value;
        const eta = document.getElementById('commentETA').value;
        
        const tabData = STATE.allData[STATE.currentTab];
        
        // Update or add comment
        const commentIndex = tabData.comments.findIndex(c => c.PRIMARY_KEY == primaryKeyValue);
        const commentData = {
            PRIMARY_KEY: primaryKeyValue,
            'ROOT CAUSE': rootCause,
            'ACTION': action,
            'ETA': eta,
            'USERNAME': STATE.currentUser,
            'UPDATED TIME': new Date().toLocaleString('en-GB', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            })
        };
        
        if (commentIndex >= 0) {
            tabData.comments[commentIndex] = commentData;
        } else {
            tabData.comments.push(commentData);
        }
        
        // Save to Excel
        const ws = XLSX.utils.json_to_sheet(tabData.comments);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(CONFIG.commentDBFile, wbout);
        
        showMessage('Comment saved successfully!', 'success');
        closeCommentModal();
        renderTable();
        
    } catch (error) {
        console.error('Error saving comment:', error);
        showMessage('Failed to save comment: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}


// ===== WRITE-BACK (ADD/EDIT/DELETE) =====
function openEditRowModal(rowIndex) {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const row = tabData.filteredData[rowIndex];
    const columns = STATE.columnOrders[STATE.currentTab] || Object.keys(row);
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = rowIndex;
    
    columns.forEach(column => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        formGroup.innerHTML = `
            <label>${column}</label>
            <input type="text" class="form-control" data-column="${column}" value="${row[column] || ''}">
        `;
        container.appendChild(formGroup);
    });
    
    document.getElementById('editRowTitle').textContent = 'Edit Row';
    document.getElementById('deleteRowButton').style.display = 'inline-flex';
    document.getElementById('editRowModal').classList.add('show');
}

function openAddRowModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const columns = STATE.columnOrders[STATE.currentTab] || 
                   (tabData.rawData.length > 0 ? Object.keys(tabData.rawData[0]) : []);
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = '-1'; // -1 indicates new row
    
    columns.forEach(column => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        formGroup.innerHTML = `
            <label>${column}</label>
            <input type="text" class="form-control" data-column="${column}" value="">
        `;
        container.appendChild(formGroup);
    });
    
    document.getElementById('editRowTitle').textContent = 'Add New Row';
    document.getElementById('deleteRowButton').style.display = 'none';
    document.getElementById('editRowModal').classList.add('show');
}

function closeEditRowModal() {
    document.getElementById('editRowModal').classList.remove('show');
}

async function saveRow() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        
        // Gather form data
        const newRow = {};
        container.querySelectorAll('.form-control').forEach(input => {
            newRow[input.dataset.column] = input.value;
        });
        
        let action = '';
        if (rowIndex === -1) {
            // Add new row
            tabData.rawData.push(newRow);
            action = 'ADD';
        } else {
            // Update existing row
            const oldRow = {...tabData.filteredData[rowIndex]};
            tabData.rawData = tabData.rawData.map(r => 
                r[CONFIG.primaryKey] === oldRow[CONFIG.primaryKey] ? newRow : r
            );
            action = 'EDIT';
        }
        
        // Save to Excel
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const ws = XLSX.utils.json_to_sheet(tabData.rawData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(tab.filename, wbout);
        
        // Log to audit
        await logAudit(action, STATE.currentTab, newRow[CONFIG.primaryKey], JSON.stringify(newRow));
        
        showMessage('Row saved successfully!', 'success');
        closeEditRowModal();
        
        // Reload tab data
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Error saving row:', error);
        showMessage('Failed to save row: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function deleteRow() {
    if (!confirm('Are you sure you want to delete this row?')) {
        return;
    }
    
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        
        const row = tabData.filteredData[rowIndex];
        const primaryKeyValue = row[CONFIG.primaryKey];
        
        // Remove from raw data
        tabData.rawData = tabData.rawData.filter(r => r[CONFIG.primaryKey] !== primaryKeyValue);
        
        // Save to Excel
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const ws = XLSX.utils.json_to_sheet(tabData.rawData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(tab.filename, wbout);
        
        // Log to audit
        await logAudit('DELETE', STATE.currentTab, primaryKeyValue, JSON.stringify(row));
        
        showMessage('Row deleted successfully!', 'success');
        closeEditRowModal();
        
        // Reload tab data
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Error deleting row:', error);
        showMessage('Failed to delete row: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function logAudit(action, tabName, primaryKey, details) {
    try {
        let auditData = [];
        
        // Try to load existing audit log
        try {
            const auditBuffer = await fetchAttachment(CONFIG.auditLogFile);
            const auditWorkbook = XLSX.read(auditBuffer, { type: 'array' });
            const auditSheet = auditWorkbook.Sheets[auditWorkbook.SheetNames[0]];
            auditData = XLSX.utils.sheet_to_json(auditSheet);
        } catch (error) {
            console.log('Creating new audit log file');
        }
        
        // Add new audit entry
        auditData.push({
            TIMESTAMP: new Date().toLocaleString(),
            USERNAME: STATE.currentUser,
            ACTION: action,
            TAB: tabName,
            PRIMARY_KEY: primaryKey,
            DETAILS: details
        });
        
        // Save audit log
        const ws = XLSX.utils.json_to_sheet(auditData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(CONFIG.auditLogFile, wbout);
        
    } catch (error) {
        console.error('Failed to log audit:', error);
    }
}

// ===== EXPORT =====
function exportData() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const dataToExport = tabData.filteredData;
    
    if (dataToExport.length === 0) {
        showMessage('No data to export', 'warning');
        return;
    }
    
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, STATE.currentTab);
    
    const filename = `${STATE.currentTab}_export_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
    
    showMessage('Data exported successfully!', 'success');
}

// ===== CONFIGURATION MODAL =====
function openConfigModal() {
    // Populate form with current config
    document.getElementById('configPageId').value = CONFIG.pageId || '';
    document.getElementById('configPrimaryKey').value = CONFIG.primaryKey || 'ITEM_ID';
    document.getElementById('configSumColumn').value = CONFIG.sumColumn || 'Amount';
    document.getElementById('configTheme').value = CONFIG.theme || 'auto';
    document.getElementById('configWriteBack').checked = CONFIG.writeBackEnabled || false;
    
    // Render tab configurations
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
            <div class="form-group">
                <label>Excel Filename</label>
                <input type="text" class="form-control tab-filename-input" value="${tab.filename}">
            </div>
            <div class="form-group">
                <label>Display Name</label>
                <input type="text" class="form-control tab-name-input" value="${tab.displayName}">
            </div>
            <button class="remove-tab-btn" onclick="removeTabConfig(${index})">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"></line>
                    <line x1="6" y1="6" x2="18" y2="18"></line>
                </svg>
            </button>
        `;
        container.appendChild(tabConfig);
    });
}

function addTabConfig() {
    STATE.tabs.push({
        key: `TAB_${STATE.tabs.length + 1}`,
        name: '',
        filename: '',
        displayName: 'New Tab'
    });
    renderTabsConfig();
}

function removeTabConfig(index) {
    if (STATE.tabs.length === 1) {
        showMessage('Cannot remove the last tab', 'warning');
        return;
    }
    STATE.tabs.splice(index, 1);
    renderTabsConfig();
}

// ===== COLUMN ORDER MODAL =====
function openColumnOrderModal(tabName) {
    const tabData = STATE.allData[tabName];
    if (!tabData || tabData.rawData.length === 0) {
        showMessage('No data loaded for this tab', 'warning');
        return;
    }
    
    const allColumns = Object.keys(tabData.rawData[0]);
    const selectedColumns = STATE.columnOrders[tabName] || allColumns;
    const availableColumns = allColumns.filter(col => !selectedColumns.includes(col));
    
    document.getElementById('columnOrderTabName').textContent = tabName;
    
    // Render available columns
    const availableList = document.getElementById('availableColumnsList');
    availableList.innerHTML = '';
    availableColumns.forEach(col => {
        const item = createColumnItem(col, false);
        availableList.appendChild(item);
    });
    
    // Render selected columns
    const selectedList = document.getElementById('selectedColumnsList');
    selectedList.innerHTML = '';
    selectedColumns.forEach(col => {
        const item = createColumnItem(col, true);
        selectedList.appendChild(item);
    });
    
    // Setup drag and drop
    setupDragAndDrop();
    
    document.getElementById('columnOrderModal').classList.add('show');
}

function closeColumnOrderModal() {
    document.getElementById('columnOrderModal').classList.remove('show');
}

function createColumnItem(columnName, isSelected) {
    const item = document.createElement('div');
    item.className = 'column-item';
    item.draggable = true;
    item.dataset.column = columnName;
    
    // Check if it's a custom column from commentdb
    const isCustom = ['ROOT CAUSE', 'ACTION', 'ETA', 'USERNAME', 'UPDATED TIME'].includes(columnName);
    
    item.innerHTML = `
        <svg class="column-handle" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="4" y1="6" x2="20" y2="6"></line>
            <line x1="4" y1="12" x2="20" y2="12"></line>
            <line x1="4" y1="18" x2="20" y2="18"></line>
        </svg>
        <span class="column-name">${columnName}</span>
        ${isCustom ? '<span class="column-badge">Custom</span>' : ''}
    `;
    
    return item;
}

function setupDragAndDrop() {
    const columnItems = document.querySelectorAll('.column-item');
    const columnLists = document.querySelectorAll('.column-list');
    
    columnItems.forEach(item => {
        item.addEventListener('dragstart', handleDragStart);
        item.addEventListener('dragend', handleDragEnd);
    });
    
    columnLists.forEach(list => {
        list.addEventListener('dragover', handleDragOver);
        list.addEventListener('drop', handleDrop);
        list.addEventListener('dragleave', handleDragLeave);
    });
}

function handleDragStart(e) {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/html', this.innerHTML);
    e.dataTransfer.setData('column', this.dataset.column);
    this.classList.add('dragging');
}

function handleDragEnd(e) {
    this.classList.remove('dragging');
    document.querySelectorAll('.column-list').forEach(list => {
        list.classList.remove('drag-over');
    });
}

function handleDragOver(e) {
    if (e.preventDefault) {
        e.preventDefault();
    }
    e.dataTransfer.dropEffect = 'move';
    this.classList.add('drag-over');
    return false;
}

function handleDragLeave(e) {
    this.classList.remove('drag-over');
}

function handleDrop(e) {
    if (e.stopPropagation) {
        e.stopPropagation();
    }
    
    const column = e.dataTransfer.getData('column');
    const draggingElement = document.querySelector('.dragging');
    
    if (draggingElement) {
        this.appendChild(draggingElement);
    }
    
    this.classList.remove('drag-over');
    return false;
}

function saveColumnOrder() {
    const tabName = document.getElementById('columnOrderTabName').textContent;
    const selectedList = document.getElementById('selectedColumnsList');
    const selectedColumns = Array.from(selectedList.querySelectorAll('.column-item'))
        .map(item => item.dataset.column);
    
    STATE.columnOrders[tabName] = selectedColumns;
    
    showMessage('Column order saved!', 'success');
    closeColumnOrderModal();
    renderTable();
}

// ===== UTILITIES =====
function formatCurrency(value) {
    const num = parseFloat(value) || 0;
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 2
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

function showMessage(message, type = 'info') {
    // Simple alert for now - can be enhanced with toast notifications
    const icon = type === 'success' ? '✓' : type === 'error' ? '✗' : 'ℹ';
    alert(`${icon} ${message}`);
}

// ===== EVENT LISTENERS =====
function setupEventListeners() {
    // Config button
    document.getElementById('configButton').addEventListener('click', openConfigModal);
    document.getElementById('saveConfigButton').addEventListener('click', saveConfiguration);
    document.getElementById('addTabButton').addEventListener('click', addTabConfig);
    
    // Filter
    document.getElementById('filterButton').addEventListener('click', openFilterModal);
    document.getElementById('applyFiltersButton').addEventListener('click', applyFiltersFromModal);
    
    // Search
    document.getElementById('searchInput').addEventListener('input', function() {
        STATE.searchQuery = this.value.toLowerCase();
        applyFilters();
        renderTable();
        updateSummary();
    });
    
    // Refresh
    document.getElementById('refreshButton').addEventListener('click', async function() {
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
    });
    
    // Export
    document.getElementById('exportButton').addEventListener('click', exportData);
    
    // Selection
    document.getElementById('selectAllButton').addEventListener('click', selectAllRows);
    document.getElementById('unselectAllButton').addEventListener('click', unselectAllRows);
    
    // Pagination
    document.getElementById('firstPageBtn').addEventListener('click', () => goToPage(1));
    document.getElementById('prevPageBtn').addEventListener('click', () => goToPage(STATE.currentPage - 1));
    document.getElementById('nextPageBtn').addEventListener('click', () => goToPage(STATE.currentPage + 1));
    document.getElementById('lastPageBtn').addEventListener('click', () => {
        const tabData = STATE.allData[STATE.currentTab];
        const totalPages = Math.ceil(tabData.filteredData.length / STATE.pageSize);
        goToPage(totalPages);
    });
    
    document.getElementById('pageSizeSelect').addEventListener('change', function() {
        STATE.pageSize = this.value;
        STATE.currentPage = 1;
        renderTable();
    });
    
    // Write-back
    document.getElementById('addRowButton').addEventListener('click', openAddRowModal);
    document.getElementById('saveRowButton').addEventListener('click', saveRow);
    document.getElementById('deleteRowButton').addEventListener('click', deleteRow);
    
    // Comments
    document.getElementById('saveCommentButton').addEventListener('click', saveComment);
    
    // Column order
    document.getElementById('saveColumnOrderButton').addEventListener('click', saveColumnOrder);
    
    // Close modals on outside click
    window.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal')) {
            e.target.classList.remove('show');
        }
    });
}

// ===== GLOBAL FUNCTIONS (called from HTML) =====
window.openCommentModal = openCommentModal;
window.closeCommentModal = closeCommentModal;
window.openEditRowModal = openEditRowModal;
window.closeEditRowModal = closeEditRowModal;
window.openConfigModal = openConfigModal;
window.closeConfigModal = closeConfigModal;
window.openColumnOrderModal = openColumnOrderModal;
window.closeColumnOrderModal = closeColumnOrderModal;
window.openFilterModal = openFilterModal;
window.closeFilterModal = closeFilterModal;
window.clearFilters = clearFilters;
window.removeTabConfig = removeTabConfig;
window.addTabConfig = addTabConfig;

console.log('Worksheet Tool loaded successfully!');

