// WORKSHEET TOOL - Complete JavaScript
// ============================================

// ===== CONFIGURATION =====
const CONFIG = {
    confluenceBaseUrl: 'https://your-domain.atlassian.net/wiki',
    pageId: '',
    worksheetUIFile: 'worksheetui.xlsx',
    commentDBFile: 'commentdb.xlsx',
    auditLogFile: 'audit_log.xlsx',
    worksheetName: 'Worksheet',
    primaryKey: 'ITEM_ID',
    sumColumn: 'Amount',
    theme: 'auto',
    writeBackEnabled: false
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
    searchQuery: '',
    currentUser: 'Loading...',
    columnOrders: {},
    commentColumns: []
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', async function() {
    console.log('Worksheet Tool initializing...');
    await init();
});

async function init() {
    try {
        showLoading(true);
        await getCurrentUser();
        
        // Try to load configuration
        try {
            await loadConfiguration();
        } catch (error) {
            console.warn('Config not found, opening config modal');
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
        console.error('Initialization error:', error);
        showMessage('Failed to initialize: ' + error.message, 'error');
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
            STATE.currentUser = user.displayName || user.username || 'Unknown';
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
    
    return await fileResponse.arrayBuffer();
}

async function uploadAttachment(filename, fileContent) {
    const attachmentsUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    const attachmentsResponse = await fetch(attachmentsUrl, { credentials: 'include' });
    
    if (!attachmentsResponse.ok) {
        throw new Error(`Failed to fetch attachments: ${attachmentsResponse.status}`);
    }
    
    const attachmentsData = await attachmentsResponse.json();
    const existingAttachment = attachmentsData.results.find(att => att.title === filename);
    
    const formData = new FormData();
    const blob = new Blob([fileContent], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
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
        const errorText = await response.text();
        throw new Error(`Upload failed: ${response.status} - ${errorText}`);
    }
    
    return await response.json();
}

// ===== CONFIGURATION =====
async function loadConfiguration() {
    const configBuffer = await fetchAttachment(CONFIG.worksheetUIFile);
    const workbook = XLSX.read(configBuffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const configData = XLSX.utils.sheet_to_json(sheet, { header: ['Fields', 'Values'] });
    
    STATE.tabs = [];
    
    configData.forEach(row => {
        const field = row.Fields;
        const value = row.Values;
        
        if (!field || value === undefined || value === null || value === '') return;
        
        if (field === 'WORKSHEET_NAME') {
            CONFIG.worksheetName = value.toString();
        } else if (field === 'CONFLUENCE_BASE_URL') {
            CONFIG.confluenceBaseUrl = value.toString();
        } else if (field === 'PAGEID') {
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
        } else if (field.match(/TAB_FILE_\d+_NAME$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            const tab = STATE.tabs.find(t => t.key === tabKey);
            if (tab) tab.displayName = value;
        } else if (field.match(/TAB_FILE_\d+_COLUMN_ORDER$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            const tab = STATE.tabs.find(t => t.key === tabKey);
            if (tab && value) {
                STATE.columnOrders[tab.name] = value.split(',').map(c => c.trim());
            }
        }
    });
    
    renderTabs();
    
    if (CONFIG.writeBackEnabled) {
        document.getElementById('editRowButton').style.display = 'inline-flex';
        document.getElementById('addRowButton').style.display = 'inline-flex';
        document.getElementById('deleteRowsButton').style.display = 'inline-flex';
    }
}

async function saveConfiguration() {
    try {
        showLoading(true);
        
        CONFIG.worksheetName = document.getElementById('configWorksheetName').value.trim();
        CONFIG.confluenceBaseUrl = document.getElementById('configBaseUrl').value.trim();
        CONFIG.pageId = document.getElementById('configPageId').value.trim();
        CONFIG.primaryKey = document.getElementById('configPrimaryKey').value.trim();
        CONFIG.sumColumn = document.getElementById('configSumColumn').value.trim();
        CONFIG.theme = document.getElementById('configTheme').value;
        CONFIG.writeBackEnabled = document.getElementById('configWriteBack').checked;
        
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
        
        const configRows = [
            ['Fields', 'Values'],
            ['WORKSHEET_NAME', CONFIG.worksheetName],
            ['CONFLUENCE_BASE_URL', CONFIG.confluenceBaseUrl],
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
        
        showMessage('Configuration saved!', 'success');
        closeConfigModal();
        await init();
        
    } catch (error) {
        console.error('Save config error:', error);
        showMessage('Failed to save: ' + error.message, 'error');
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
        
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tabName === tabName);
        });
        
        if (!STATE.allData[tabName]) {
            await loadTabData(tabName);
        }
        
        applyFilters();
        renderTable();
        updateSummary();
        updateButtonStates();
        
        // Show customize columns button
        document.getElementById('customizeColumnsButton').style.display = 'inline-flex';
        
        showLoading(false);
    } catch (error) {
        console.error('Switch tab error:', error);
        showMessage('Failed to load tab: ' + error.message, 'error');
        showLoading(false);
    }
}

async function loadTabData(tabName) {
    const tab = STATE.tabs.find(t => t.name === tabName);
    if (!tab) return;
    
    const buffer = await fetchAttachment(tab.filename);
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);
    
    // Load comments
    let comments = [];
    let commentColumns = [];
    try {
        const commentBuffer = await fetchAttachment(CONFIG.commentDBFile);
        const commentWorkbook = XLSX.read(commentBuffer, { type: 'array' });
        const commentSheet = commentWorkbook.Sheets[commentWorkbook.SheetNames[0]];
        comments = XLSX.utils.sheet_to_json(commentSheet);
        
        // Get comment columns (all columns except PRIMARY_KEY)
        if (comments.length > 0) {
            commentColumns = Object.keys(comments[0]).filter(col => col !== 'PRIMARY_KEY');
        } else {
            // Get from sheet headers even if no data
            const range = XLSX.utils.decode_range(commentSheet['!ref']);
            commentColumns = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const address = XLSX.utils.encode_col(C) + '1';
                const cell = commentSheet[address];
                if (cell && cell.v !== 'PRIMARY_KEY') {
                    commentColumns.push(cell.v);
                }
            }
        }
        
        STATE.commentColumns = commentColumns;
    } catch (error) {
        console.warn('No comments file, continuing without comments');
        STATE.commentColumns = [];
    }
    
    STATE.allData[tabName] = {
        rawData: data,
        filteredData: [...data],
        comments: comments
    };
}

// ===== TABLE RENDERING =====
function renderTable() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.filteredData;
    const comments = tabData.comments;
    
    // Get columns: comment columns first, then data columns
    let dataColumns = STATE.columnOrders[STATE.currentTab];
    if (!dataColumns || dataColumns.length === 0) {
        dataColumns = data.length > 0 ? Object.keys(data[0]) : [];
    }
    
    const allColumns = [...STATE.commentColumns, ...dataColumns];
    
    // Render header
    const thead = document.getElementById('tableHeader');
    thead.innerHTML = `
        <tr>
            <th><input type="checkbox" id="headerCheckbox" class="row-checkbox"></th>
            ${allColumns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;
    
    document.getElementById('headerCheckbox').addEventListener('change', function() {
        if (this.checked) {
            selectAllRows();
        } else {
            unselectAllRows();
        }
    });
    
    // Pagination
    const totalRows = data.length;
    const pageSize = parseInt(STATE.pageSize);
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
        
        // Checkbox
        const checkboxTd = document.createElement('td');
        checkboxTd.innerHTML = `<input type="checkbox" class="row-checkbox" ${isSelected ? 'checked' : ''}>`;
        checkboxTd.querySelector('.row-checkbox').addEventListener('change', function() {
            toggleRowSelection(globalIndex, this.checked);
        });
        tr.appendChild(checkboxTd);
        
        // Comment columns first
        STATE.commentColumns.forEach(col => {
            const td = document.createElement('td');
            const comment = comments.find(c => c.PRIMARY_KEY == primaryKeyValue);
            td.textContent = comment ? (comment[col] || '') : '';
            tr.appendChild(td);
        });
        
        // Data columns
        dataColumns.forEach(col => {
            const td = document.createElement('td');
            const value = row[col];
            
            if (col === CONFIG.sumColumn && !isNaN(value)) {
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
    document.getElementById('nextPageBtn').disabled = STATE.currentPage === totalPages;
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
        const uniqueValues = [...new Set(data.map(row => row[column]?.toString() || ''))].sort();
        
        const isActive = STATE.activeFilters[column] && STATE.activeFilters[column].length > 0;
        
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-column-group' + (isActive ? ' active' : '');
        filterGroup.innerHTML = `
            <label>${column}</label>
            <select multiple class="filter-select" data-column="${column}">
                ${uniqueValues.map(val => {
                    const isSelected = STATE.activeFilters[column]?.includes(val);
                    return `<option value="${val}" ${isSelected ? 'selected' : ''}>${val}</option>`;
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
    updateButtonStates();
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
    updateButtonStates();
}

function unselectAllRows() {
    STATE.selectedRows.clear();
    document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = false);
    document.querySelectorAll('tbody tr').forEach(tr => tr.classList.remove('selected'));
    updateSummary();
    updateButtonStates();
}

function updateButtonStates() {
    const selectedCount = STATE.selectedRows.size;
    
    // Edit button: only enabled if exactly 1 row selected
    const editBtn = document.getElementById('editRowButton');
    if (editBtn.style.display !== 'none') {
        editBtn.disabled = selectedCount !== 1;
    }
    
    // Delete button: enabled if at least 1 row selected
    const deleteBtn = document.getElementById('deleteRowsButton');
    if (deleteBtn.style.display !== 'none') {
        deleteBtn.disabled = selectedCount === 0;
    }
    
    // Comments button: enabled if at least 1 row selected
    document.getElementById('addCommentsButton').disabled = selectedCount === 0;
}

// ===== SUMMARY =====
function updateSummary() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const filteredData = tabData.filteredData;
    const totalRows = filteredData.length;
    const selectedCount = STATE.selectedRows.size;
    
    document.getElementById('totalRows').textContent = totalRows;
    document.getElementById('selectedRows').textContent = selectedCount;
    
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
    
    document.getElementById('sumValue').textContent = formatNumber(sum);
}


// ===== COMMENTS (BULK) =====
function openCommentsModal() {
    const selectedCount = STATE.selectedRows.size;
    if (selectedCount === 0) {
        showMessage('Please select at least one row', 'warning');
        return;
    }
    
    document.getElementById('commentRowCount').textContent = selectedCount;
    
    // Build form based on comment columns
    const container = document.getElementById('commentFormContainer');
    container.innerHTML = '';
    
    if (STATE.commentColumns.length === 0) {
        container.innerHTML = '<p>No comment columns configured in commentdb.xlsx</p>';
        return;
    }
    
    STATE.commentColumns.forEach(col => {
        if (col === 'USERNAME' || col === 'UPDATED TIME') return; // Skip auto fields
        
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        formGroup.innerHTML = `
            <label>${col}</label>
            <textarea class="form-control" data-column="${col}" rows="3"></textarea>
        `;
        container.appendChild(formGroup);
    });
    
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
        
        // Gather comment data
        const commentData = {};
        document.querySelectorAll('#commentFormContainer textarea').forEach(textarea => {
            commentData[textarea.dataset.column] = textarea.value;
        });
        
        // Add username and timestamp
        commentData['USERNAME'] = STATE.currentUser;
        commentData['UPDATED TIME'] = new Date().toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
        
        // Get selected rows
        const selectedData = tabData.filteredData.filter((row, index) => STATE.selectedRows.has(index));
        
        // Update or add comments for each selected row
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
        
        // Save to Excel
        const ws = XLSX.utils.json_to_sheet(tabData.comments);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(CONFIG.commentDBFile, wbout);
        
        showMessage('Comments saved!', 'success');
        closeCommentModal();
        renderTable();
        
    } catch (error) {
        console.error('Save comments error:', error);
        showMessage('Failed to save comments: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== EDIT ROW (SINGLE) =====
function openEditRowModal() {
    if (STATE.selectedRows.size !== 1) {
        showMessage('Please select exactly one row to edit', 'warning');
        return;
    }
    
    const tabData = STATE.allData[STATE.currentTab];
    const selectedIndex = Array.from(STATE.selectedRows)[0];
    const row = tabData.filteredData[selectedIndex];
    
    const columns = STATE.columnOrders[STATE.currentTab] || Object.keys(row);
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = selectedIndex;
    
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
    document.getElementById('editRowModal').classList.add('show');
}

function closeEditRowModal() {
    document.getElementById('editRowModal').classList.remove('show');
}

async function saveEditedRow() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        
        const oldRow = tabData.filteredData[rowIndex];
        const newRow = {};
        
        container.querySelectorAll('.form-control').forEach(input => {
            newRow[input.dataset.column] = input.value;
        });
        
        // Update in raw data
        tabData.rawData = tabData.rawData.map(r => 
            r[CONFIG.primaryKey] === oldRow[CONFIG.primaryKey] ? newRow : r
        );
        
        // Save to Excel
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const ws = XLSX.utils.json_to_sheet(tabData.rawData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(tab.filename, wbout);
        await logAudit('EDIT', STATE.currentTab, newRow[CONFIG.primaryKey], JSON.stringify(newRow));
        
        showMessage('Row updated!', 'success');
        closeEditRowModal();
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Edit row error:', error);
        showMessage('Failed to save: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ADD ROW =====
function openAddRowModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const columns = STATE.columnOrders[STATE.currentTab] || 
                   (tabData.rawData.length > 0 ? Object.keys(tabData.rawData[0]) : []);
    
    const container = document.getElementById('editRowFormContainer');
    container.innerHTML = '';
    container.dataset.rowIndex = '-1';
    
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
    document.getElementById('editRowModal').classList.add('show');
}

async function saveNewRow() {
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const container = document.getElementById('editRowFormContainer');
        const rowIndex = parseInt(container.dataset.rowIndex);
        
        if (rowIndex === -1) {
            // Add new row
            const newRow = {};
            container.querySelectorAll('.form-control').forEach(input => {
                newRow[input.dataset.column] = input.value;
            });
            
            tabData.rawData.push(newRow);
            
            const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
            const ws = XLSX.utils.json_to_sheet(tabData.rawData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            
            await uploadAttachment(tab.filename, wbout);
            await logAudit('ADD', STATE.currentTab, newRow[CONFIG.primaryKey], JSON.stringify(newRow));
            
            showMessage('Row added!', 'success');
            closeEditRowModal();
            
            delete STATE.allData[STATE.currentTab];
            await switchTab(STATE.currentTab);
        } else {
            // Edit existing
            await saveEditedRow();
        }
        
    } catch (error) {
        console.error('Add row error:', error);
        showMessage('Failed to add: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ===== DELETE ROWS (BULK) =====
async function deleteSelectedRows() {
    if (STATE.selectedRows.size === 0) {
        showMessage('Please select rows to delete', 'warning');
        return;
    }
    
    if (!confirm(`Delete ${STATE.selectedRows.size} row(s)?`)) {
        return;
    }
    
    try {
        showLoading(true);
        
        const tabData = STATE.allData[STATE.currentTab];
        const selectedData = tabData.filteredData.filter((row, index) => STATE.selectedRows.has(index));
        const primaryKeysToDelete = selectedData.map(row => row[CONFIG.primaryKey]);
        
        // Remove from raw data
        tabData.rawData = tabData.rawData.filter(r => !primaryKeysToDelete.includes(r[CONFIG.primaryKey]));
        
        // Save
        const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
        const ws = XLSX.utils.json_to_sheet(tabData.rawData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(tab.filename, wbout);
        await logAudit('DELETE', STATE.currentTab, primaryKeysToDelete.join(','), 'Bulk delete');
        
        showMessage('Rows deleted!', 'success');
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Delete error:', error);
        showMessage('Failed to delete: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function logAudit(action, tabName, primaryKey, details) {
    try {
        let auditData = [];
        
        try {
            const auditBuffer = await fetchAttachment(CONFIG.auditLogFile);
            const auditWorkbook = XLSX.read(auditBuffer, { type: 'array' });
            const auditSheet = auditWorkbook.Sheets[auditWorkbook.SheetNames[0]];
            auditData = XLSX.utils.sheet_to_json(auditSheet);
        } catch (error) {
            console.log('Creating new audit log');
        }
        
        auditData.push({
            TIMESTAMP: new Date().toLocaleString(),
            USERNAME: STATE.currentUser,
            ACTION: action,
            TAB: tabName,
            PRIMARY_KEY: primaryKey,
            DETAILS: details
        });
        
        const ws = XLSX.utils.json_to_sheet(auditData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        await uploadAttachment(CONFIG.auditLogFile, wbout);
    } catch (error) {
        console.warn('Audit log failed:', error);
    }
}

// ===== EXPORT =====
function exportData() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData || tabData.filteredData.length === 0) {
        showMessage('No data to export', 'warning');
        return;
    }
    
    const ws = XLSX.utils.json_to_sheet(tabData.filteredData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, STATE.currentTab);
    
    const filename = `${STATE.currentTab}_export_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
    
    showMessage('Data exported!', 'success');
}


// ===== COLUMN CUSTOMIZATION =====
function openColumnOrderModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData || tabData.rawData.length === 0) {
        showMessage('No data loaded', 'warning');
        return;
    }
    
    const allDataColumns = Object.keys(tabData.rawData[0]);
    const allColumns = [...STATE.commentColumns, ...allDataColumns];
    
    const selectedColumns = STATE.columnOrders[STATE.currentTab] || allDataColumns;
    const availableColumns = allColumns.filter(col => !selectedColumns.includes(col));
    
    const availableList = document.getElementById('availableColumnsList');
    availableList.innerHTML = '';
    availableColumns.forEach(col => {
        const item = createColumnItem(col);
        availableList.appendChild(item);
    });
    
    const selectedList = document.getElementById('selectedColumnsList');
    selectedList.innerHTML = '';
    selectedColumns.forEach(col => {
        const item = createColumnItem(col);
        selectedList.appendChild(item);
    });
    
    setupDragAndDrop();
    document.getElementById('columnOrderModal').classList.add('show');
}

function closeColumnOrderModal() {
    document.getElementById('columnOrderModal').classList.remove('show');
}

function createColumnItem(columnName) {
    const item = document.createElement('div');
    item.className = 'column-item';
    item.draggable = true;
    item.dataset.column = columnName;
    
    const isComment = STATE.commentColumns.includes(columnName);
    
    item.innerHTML = `
        <svg class="column-handle" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="4" y1="6" x2="20" y2="6"></line>
            <line x1="4" y1="12" x2="20" y2="12"></line>
            <line x1="4" y1="18" x2="20" y2="18"></line>
        </svg>
        <span style="flex: 1;">${columnName}</span>
        ${isComment ? '<span style="padding: 2px 8px; background: var(--success); color: white; font-size: 11px; border-radius: 10px;">Comment</span>' : ''}
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
    });
}

function handleDragStart(e) {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('column', this.dataset.column);
    this.classList.add('dragging');
}

function handleDragEnd(e) {
    this.classList.remove('dragging');
}

function handleDragOver(e) {
    if (e.preventDefault) {
        e.preventDefault();
    }
    e.dataTransfer.dropEffect = 'move';
    return false;
}

function handleDrop(e) {
    if (e.stopPropagation) {
        e.stopPropagation();
    }
    
    const draggingElement = document.querySelector('.dragging');
    if (draggingElement) {
        this.appendChild(draggingElement);
    }
    
    return false;
}

async function saveColumnOrder() {
    const selectedList = document.getElementById('selectedColumnsList');
    const selectedColumns = Array.from(selectedList.querySelectorAll('.column-item'))
        .map(item => item.dataset.column);
    
    // Filter out comment columns - only save data columns
    const dataColumns = selectedColumns.filter(col => !STATE.commentColumns.includes(col));
    
    STATE.columnOrders[STATE.currentTab] = dataColumns;
    
    // Save to config
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    if (tab) {
        try {
            showLoading(true);
            
            // Reload config, update this tab's column order, and save
            const configBuffer = await fetchAttachment(CONFIG.worksheetUIFile);
            const workbook = XLSX.read(configBuffer, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const configData = XLSX.utils.sheet_to_json(sheet, { header: ['Fields', 'Values'] });
            
            // Find or add column order row
            const tabIndex = STATE.tabs.findIndex(t => t.name === STATE.currentTab);
            const fieldName = `TAB_FILE_${tabIndex + 1}_COLUMN_ORDER`;
            
            const existingRow = configData.find(r => r.Fields === fieldName);
            if (existingRow) {
                existingRow.Values = dataColumns.join(',');
            } else {
                configData.push({ Fields: fieldName, Values: dataColumns.join(',') });
            }
            
            // Save
            const ws = XLSX.utils.json_to_sheet(configData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            
            await uploadAttachment(CONFIG.worksheetUIFile, wbout);
            
            showMessage('Column order saved!', 'success');
            closeColumnOrderModal();
            renderTable();
            
        } catch (error) {
            console.error('Save column order error:', error);
            showMessage('Failed to save: ' + error.message, 'error');
        } finally {
            showLoading(false);
        }
    }
}

// ===== CONFIGURATION MODAL =====
function openConfigModal() {
    document.getElementById('configWorksheetName').value = CONFIG.worksheetName || 'Worksheet';
    document.getElementById('configBaseUrl').value = CONFIG.confluenceBaseUrl || '';
    document.getElementById('configPageId').value = CONFIG.pageId || '';
    document.getElementById('configPrimaryKey').value = CONFIG.primaryKey || 'ITEM_ID';
    document.getElementById('configSumColumn').value = CONFIG.sumColumn || 'Amount';
    document.getElementById('configTheme').value = CONFIG.theme || 'auto';
    document.getElementById('configWriteBack').checked = CONFIG.writeBackEnabled || false;
    
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
            <button class="btn btn-danger btn-sm" onclick="removeTabConfig(${index})">×</button>
        `;
        container.appendChild(tabConfig);
    });
}

function addTabConfig() {
    if (STATE.tabs.length >= 5) {
        showMessage('Maximum 5 tabs allowed', 'warning');
        return;
    }
    
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

function showMessage(message, type = 'info') {
    const icon = type === 'success' ? '✓' : type === 'error' ? '✗' : type === 'warning' ? '⚠' : 'ℹ';
    alert(`${icon} ${message}`);
}

// ===== EVENT LISTENERS =====
function setupEventListeners() {
    // Config
    document.getElementById('configButton').addEventListener('click', openConfigModal);
    document.getElementById('saveConfigButton').addEventListener('click', saveConfiguration);
    document.getElementById('addTabButton').addEventListener('click', addTabConfig);
    
    // Filter
    document.getElementById('filterButton').addEventListener('click', openFilterModal);
    document.getElementById('applyFiltersButton').addEventListener('click', applyFiltersFromModal);
    
    // Columns
    document.getElementById('customizeColumnsButton').addEventListener('click', openColumnOrderModal);
    document.getElementById('saveColumnOrderButton').addEventListener('click', saveColumnOrder);
    
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
    
    // Comments
    document.getElementById('addCommentsButton').addEventListener('click', openCommentsModal);
    document.getElementById('saveCommentButton').addEventListener('click', saveComments);
    
    // Edit/Delete/Add
    document.getElementById('editRowButton').addEventListener('click', openEditRowModal);
    document.getElementById('addRowButton').addEventListener('click', openAddRowModal);
    document.getElementById('deleteRowsButton').addEventListener('click', deleteSelectedRows);
    document.getElementById('saveRowButton').addEventListener('click', saveNewRow);
    
    // Close modals on outside click
    window.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal')) {
            e.target.classList.remove('show');
        }
    });
}

// ===== GLOBAL FUNCTIONS =====
window.closeConfigModal = closeConfigModal;
window.closeColumnOrderModal = closeColumnOrderModal;
window.closeFilterModal = closeFilterModal;
window.clearFilters = clearFilters;
window.closeCommentModal = closeCommentModal;
window.closeEditRowModal = closeEditRowModal;
window.removeTabConfig = removeTabConfig;

console.log('Worksheet Tool loaded!');
