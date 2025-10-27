// WORKSHEET TOOL - Full CSV Version
// ==========================================

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
    tabFilters: {},
    tabSearch: {},
    currentUser: 'Loading...',
    commentColumns: []
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', async function() {
    console.log('Worksheet Tool (CSV Version) initializing...');
    await init();
});

async function init() {
    try {
        showLoading(true);
        await getCurrentUser();
        
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
        console.error('Init error:', error);
        showMessage('Failed to initialize: ' + error.message);
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
        throw new Error(`Failed to fetch attachments: ${attachmentsResponse.status}`);
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
        const errorText = await response.text();
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
                    columnOrder: []
                });
            }
        } else if (field.match(/^TAB_FILE_\d+_NAME$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            const tab = STATE.tabs.find(t => t.key === tabKey);
            if (tab) tab.filename = value;
        } else if (field.match(/^TAB_FILE_\d+_SUM_COLUMN$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            const tab = STATE.tabs.find(t => t.key === tabKey);
            if (tab) tab.sumColumn = value;
        } else if (field.match(/^TAB_FILE_\d+_WRITE_BACK$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            const tab = STATE.tabs.find(t => t.key === tabKey);
            if (tab) tab.writeBack = value.toUpperCase() === 'TRUE';
        } else if (field.match(/^TAB_FILE_\d+_COLUMN_ORDER$/)) {
            const tabNum = field.split('_')[2];
            const tabKey = `TAB_${tabNum}`;
            const tab = STATE.tabs.find(t => t.key === tabKey);
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
            const columnOrder = item.querySelector('.tab-columns-input').value.trim();
            
            if (displayName && filename) {
                STATE.tabs.push({
                    key: `TAB_${index + 1}`,
                    name: filename.replace('.csv', ''),
                    displayName: displayName,
                    filename: filename,
                    sumColumn: sumColumn,
                    writeBack: writeBack,
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
            if (tab.columnOrder.length > 0) {
                configRows.push({ Fields: `TAB_FILE_${index + 1}_COLUMN_ORDER`, Values: tab.columnOrder.join(',') });
            }
        });
        
        const csvContent = CSV.generate(configRows);
        await uploadAttachment(CONFIG.worksheetUIFile, csvContent);
        
        showMessage('Configuration saved!');
        closeConfigModal();
        await init();
        
    } catch (error) {
        console.error('Save config error:', error);
        showMessage('Failed to save: ' + error.message);
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
        
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tabName === tabName);
        });
        
        if (!STATE.allData[tabName]) {
            await loadTabData(tabName);
        }
        
        // Load tab-specific filters and search
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
        showMessage('Failed to load tab: ' + error.message);
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
    try {
        const commentText = await fetchAttachment(CONFIG.commentDBFile);
        comments = CSV.parse(commentText);
        
        if (comments.length > 0) {
            commentColumns = Object.keys(comments[0]).filter(col => col !== 'PRIMARY_KEY');
        }
        
        STATE.commentColumns = commentColumns;
    } catch (error) {
        console.warn('No comments file');
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
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const data = tabData.filteredData;
    const comments = tabData.comments;
    
    let dataColumns = tab.columnOrder;
    if (!dataColumns || dataColumns.length === 0) {
        dataColumns = data.length > 0 ? Object.keys(data[0]) : [];
    }
    
    const allColumns = [...STATE.commentColumns, ...dataColumns];
    
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
        
        STATE.commentColumns.forEach(col => {
            const td = document.createElement('td');
            const comment = comments.find(c => c.PRIMARY_KEY == primaryKeyValue);
            td.textContent = comment ? (comment[col] || '') : '';
            tr.appendChild(td);
        });
        
        dataColumns.forEach(col => {
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


// ===== FILTERING (PER TAB) =====
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
    
    // Save filters and search for this tab
    STATE.tabFilters[STATE.currentTab] = STATE.activeFilters;
    STATE.tabSearch[STATE.currentTab] = STATE.searchQuery;
}

function openFilterModal() {
    const tabData = STATE.allData[STATE.currentTab];
    if (!tabData) return;
    
    const data = tabData.rawData;
    if (data.length === 0) return;
    
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const columns = tab.columnOrder.length > 0 ? tab.columnOrder : Object.keys(data[0]);
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
    
    // Show/hide write-back buttons based on tab setting
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
    document.getElementById('addCommentsButton').disabled = selectedCount === 0;
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
    
    document.getElementById('sumValue').textContent = formatNumber(sum);
}

// ===== COMMENTS (BULK) =====
function openCommentsModal() {
    const selectedCount = STATE.selectedRows.size;
    if (selectedCount === 0) {
        showMessage('Please select at least one row');
        return;
    }
    
    document.getElementById('commentRowCount').textContent = selectedCount;
    
    const container = document.getElementById('commentFormContainer');
    container.innerHTML = '';
    
    if (STATE.commentColumns.length === 0) {
        container.innerHTML = '<p>No comment columns configured in commentdb.csv</p>';
        return;
    }
    
    STATE.commentColumns.forEach(col => {
        if (col === 'USERNAME' || col === 'UPDATED TIME') return;
        
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
        
        const commentData = {};
        document.querySelectorAll('#commentFormContainer textarea').forEach(textarea => {
            commentData[textarea.dataset.column] = textarea.value;
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
        
        const csvContent = CSV.generate(tabData.comments);
        await uploadAttachment(CONFIG.commentDBFile, csvContent);
        
        showMessage('Comments saved!');
        closeCommentModal();
        renderTable();
        
    } catch (error) {
        console.error('Save comments error:', error);
        showMessage('Failed to save: ' + error.message);
    } finally {
        showLoading(false);
    }
}


// ===== EDIT ROW =====
function openEditRowModal() {
    if (STATE.selectedRows.size !== 1) {
        showMessage('Please select exactly one row');
        return;
    }
    
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    const selectedIndex = Array.from(STATE.selectedRows)[0];
    const row = tabData.filteredData[selectedIndex];
    
    const columns = tab.columnOrder.length > 0 ? tab.columnOrder : Object.keys(row);
    
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
    
    document.getElementById('editRowTitle').textContent = '✏️ Edit Row';
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
        
        const oldRow = tabData.filteredData[rowIndex];
        const newRow = {};
        
        container.querySelectorAll('.form-control').forEach(input => {
            newRow[input.dataset.column] = input.value;
        });
        
        tabData.rawData = tabData.rawData.map(r => 
            r[CONFIG.primaryKey] === oldRow[CONFIG.primaryKey] ? newRow : r
        );
        
        const csvContent = CSV.generate(tabData.rawData);
        await uploadAttachment(tab.filename, csvContent);
        await logAudit('EDIT', STATE.currentTab, newRow[CONFIG.primaryKey]);
        
        showMessage('Row updated!');
        closeEditRowModal();
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Edit error:', error);
        showMessage('Failed to save: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// ===== ADD ROW =====
function openAddRowModal() {
    const tabData = STATE.allData[STATE.currentTab];
    const tab = STATE.tabs.find(t => t.name === STATE.currentTab);
    if (!tabData) return;
    
    const columns = tab.columnOrder.length > 0 ? tab.columnOrder : 
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
    
    document.getElementById('editRowTitle').textContent = '➕ Add New Row';
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
            
            showMessage('Row added!');
            closeEditRowModal();
            
            delete STATE.allData[STATE.currentTab];
            await switchTab(STATE.currentTab);
        } else {
            await saveEditedRow();
        }
        
    } catch (error) {
        console.error('Add error:', error);
        showMessage('Failed to add: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// ===== DELETE ROWS =====
async function deleteSelectedRows() {
    if (STATE.selectedRows.size === 0) {
        showMessage('Please select rows to delete');
        return;
    }
    
    if (!confirm(`Delete ${STATE.selectedRows.size} row(s)?`)) {
        return;
    }
    
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
        
        showMessage('Rows deleted!');
        
        delete STATE.allData[STATE.currentTab];
        await switchTab(STATE.currentTab);
        
    } catch (error) {
        console.error('Delete error:', error);
        showMessage('Failed to delete: ' + error.message);
    } finally {
        showLoading(false);
    }
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
        showMessage('No data to export');
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
    
    showMessage('Data exported!');
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
                <label>Column Order (comma-separated)</label>
                <input type="text" class="form-control tab-columns-input" value="${tab.columnOrder.join(',')}" placeholder="ID,Name,Amount,Date">
            </div>
            <div class="form-group">
                <label>
                    <input type="checkbox" class="tab-writeback-input" ${tab.writeBack ? 'checked' : ''}>
                    Enable Write-Back (Add/Edit/Delete)
                </label>
            </div>
        `;
        container.appendChild(tabConfig);
    });
}

function addTabConfig() {
    if (STATE.tabs.length >= 5) {
        showMessage('Maximum 5 tabs allowed');
        return;
    }
    
    STATE.tabs.push({
        key: `TAB_${STATE.tabs.length + 1}`,
        name: '',
        displayName: 'New Tab',
        filename: '',
        sumColumn: '',
        writeBack: false,
        columnOrder: []
    });
    renderTabsConfig();
}

function removeTabConfig(index) {
    if (STATE.tabs.length === 1) {
        showMessage('Cannot remove the last tab');
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

function showMessage(message) {
    alert(message);
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
window.removeTabConfig = removeTabConfig;

console.log('Worksheet Tool (CSV Version) loaded!');
