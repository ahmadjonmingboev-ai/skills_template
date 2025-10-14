// ===== CONFIGURATION =====
const CONFIG = {
    confluenceBaseUrl: 'https://your-confluence-instance.atlassian.net', // UPDATE THIS
    pageId: 'YOUR_PAGE_ID', // UPDATE THIS
    mappingFileName: 'Mapping.xlsx',
    historyFileName: 'mapping_history.xlsx'
};

// ===== STATE MANAGEMENT =====
let mappingData = [];
let filteredData = [];
let historyData = [];
let currentPage = 1;
let rowsPerPage = 50;
let sortColumn = null;
let sortDirection = 'asc';
let currentUser = 'Unknown User';
let tableColumns = [];

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', function() {
    getCurrentUser();
    loadData();
    setupGeneratorListeners();
});

// ===== NOTIFICATION SYSTEM =====
function showNotification(title, message, type = 'info') {
    const container = document.getElementById('notificationContainer');
    const notificationId = 'notification-' + Date.now();
    
    const iconSvg = {
        success: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>',
        error: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
        info: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>',
        warning: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>'
    };

    const notification = document.createElement('div');
    notification.id = notificationId;
    notification.className = `notification ${type}`;
    notification.innerHTML = `
        ${iconSvg[type]}
        <div class="notification-content">
            <div class="notification-title">${title}</div>
            <div class="notification-message">${message}</div>
        </div>
    `;

    container.appendChild(notification);

    // Auto remove after 5 seconds
    setTimeout(() => {
        notification.classList.add('fade-out');
        setTimeout(() => {
            notification.remove();
        }, 300);
    }, 5000);
}

// ===== GET CURRENT USER =====
async function getCurrentUser() {
    try {
        const response = await fetch(`${CONFIG.confluenceBaseUrl}/rest/api/user/current`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
            },
            credentials: 'include'
        });

        if (response.ok) {
            const userData = await response.json();
            currentUser = userData.displayName || userData.username || 'Unknown User';
            document.getElementById('currentUser').textContent = currentUser;
        } else {
            console.log('Could not fetch current user');
            document.getElementById('currentUser').textContent = 'Guest User';
        }
    } catch (error) {
        console.error('Error fetching current user:', error);
        document.getElementById('currentUser').textContent = 'Guest User';
    }
}

// ===== TAB SWITCHING =====
function switchTab(tabName) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Remove active class from all buttons
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });

    // Show selected tab
    if (tabName === 'table') {
        document.getElementById('table-tab').classList.add('active');
        document.querySelectorAll('.tab-button')[0].classList.add('active');
    } else if (tabName === 'generator') {
        document.getElementById('generator-tab').classList.add('active');
        document.querySelectorAll('.tab-button')[1].classList.add('active');
        updateOwnersList();
    }
}

// ===== DATA LOADING =====
async function loadData() {
    try {
        showLoadingState();
        
        // Fetch mapping data from Confluence
        const mappingFile = await fetchFileFromConfluence(CONFIG.mappingFileName);
        const workbook = XLSX.read(mappingFile, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        mappingData = XLSX.utils.sheet_to_json(worksheet);

        // Determine columns from the first row
        if (mappingData.length > 0) {
            tableColumns = Object.keys(mappingData[0]);
            generateTableHeaders();
            generateFormFields();
        }

        // Try to fetch history data
        try {
            const historyFile = await fetchFileFromConfluence(CONFIG.historyFileName);
            const historyWorkbook = XLSX.read(historyFile, { type: 'array' });
            const historyWorksheet = historyWorkbook.Sheets[historyWorkbook.SheetNames[0]];
            historyData = XLSX.utils.sheet_to_json(historyWorksheet);
        } catch (error) {
            console.log('History file not found, initializing empty history');
            historyData = [];
        }

        populateFilters();
        filterTable();
        
        showNotification('Success', 'Data loaded successfully!', 'success');
    } catch (error) {
        console.error('Error loading data:', error);
        showNotification('Error', 'Failed to load data. Please check your Confluence configuration.', 'error');
    }
}

// ===== DYNAMIC TABLE HEADERS =====
function generateTableHeaders() {
    const thead = document.getElementById('tableHeader');
    let headerHTML = '<tr>';
    
    tableColumns.forEach(column => {
        // Format column name for display
        const displayName = column.replace(/_/g, ' ');
        headerHTML += `<th class="sortable" onclick="sortTable('${column}')">${displayName}</th>`;
    });
    
    headerHTML += '<th>Actions</th></tr>';
    thead.innerHTML = headerHTML;
}

// ===== DYNAMIC FORM FIELDS =====
function generateFormFields() {
    const formFields = document.getElementById('formFields');
    let formHTML = '<input type="hidden" id="editIndex">';
    
    // Group fields by rows (3 per row)
    const fieldsPerRow = 3;
    for (let i = 0; i < tableColumns.length; i += fieldsPerRow) {
        formHTML += '<div class="form-row">';
        
        for (let j = i; j < Math.min(i + fieldsPerRow, tableColumns.length); j++) {
            const column = tableColumns[j];
            const displayName = column.replace(/_/g, ' ');
            const fieldId = `edit${column}`;
            
            // Special handling for TLM_INSTANCE field
            if (column === 'TLM_INSTANCE' || column === 'INSTANCE') {
                formHTML += `
                    <div class="form-group">
                        <label>${displayName}: *</label>
                        <select id="${fieldId}" required>
                            <option value="">-- Select --</option>
                            <option value="CASH">CASH</option>
                            <option value="STOCK">STOCK</option>
                        </select>
                    </div>
                `;
            } else {
                const isRequired = column.includes('CODE') ? 'required' : '';
                formHTML += `
                    <div class="form-group">
                        <label>${displayName}:${isRequired ? ' *' : ''}</label>
                        <input type="text" id="${fieldId}" ${isRequired} 
                               placeholder="${displayName}">
                    </div>
                `;
            }
        }
        
        formHTML += '</div>';
    }
    
    formFields.innerHTML = formHTML;
}

async function fetchFileFromConfluence(fileName) {
    // Get attachments list
    const attachmentsUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    const response = await fetch(attachmentsUrl, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
        },
        credentials: 'include'
    });

    if (!response.ok) {
        throw new Error(`Failed to fetch attachments: ${response.statusText}`);
    }

    const data = await response.json();
    const attachment = data.results.find(att => att.title === fileName);

    if (!attachment) {
        throw new Error(`File ${fileName} not found`);
    }

    // Download the file
    const downloadUrl = `${CONFIG.confluenceBaseUrl}${attachment._links.download}`;
    const fileResponse = await fetch(downloadUrl, {
        credentials: 'include'
    });

    if (!fileResponse.ok) {
        throw new Error(`Failed to download file: ${fileResponse.statusText}`);
    }

    return await fileResponse.arrayBuffer();
}

// ===== FILTERING & SEARCH =====
function populateFilters() {
    const instanceFilter = document.getElementById('instanceFilter');
    const instanceColumn = tableColumns.find(col => 
        col === 'TLM_INSTANCE' || col === 'INSTANCE'
    ) || 'TLM_INSTANCE';
    
    const instances = [...new Set(mappingData.map(row => row[instanceColumn]))];
    
    instanceFilter.innerHTML = '<option value="">All</option>';
    instances.forEach(instance => {
        if (instance) {
            instanceFilter.innerHTML += `<option value="${instance}">${instance}</option>`;
        }
    });
}

function filterTable() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const instanceFilter = document.getElementById('instanceFilter').value;
    
    const instanceColumn = tableColumns.find(col => 
        col === 'TLM_INSTANCE' || col === 'INSTANCE'
    ) || 'TLM_INSTANCE';

    filteredData = mappingData.filter(row => {
        // Search filter
        const matchesSearch = searchTerm === '' || 
            Object.values(row).some(val => 
                String(val).toLowerCase().includes(searchTerm)
            );

        // Instance filter
        const matchesInstance = instanceFilter === '' || row[instanceColumn] === instanceFilter;

        return matchesSearch && matchesInstance;
    });

    currentPage = 1;
    renderTable();
}

// ===== SORTING =====
function sortTable(column) {
    if (sortColumn === column) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = column;
        sortDirection = 'asc';
    }

    filteredData.sort((a, b) => {
        let aVal = a[column] || '';
        let bVal = b[column] || '';
        
        if (sortDirection === 'asc') {
            return aVal > bVal ? 1 : -1;
        } else {
            return aVal < bVal ? 1 : -1;
        }
    });

    updateSortIndicators();
    renderTable();
}

function updateSortIndicators() {
    document.querySelectorAll('th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
    });

    if (sortColumn) {
        const headers = document.querySelectorAll('th.sortable');
        headers.forEach(th => {
            if (th.textContent.includes(sortColumn.replace(/_/g, ' '))) {
                th.classList.add(sortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
            }
        });
    }
}

// ===== TABLE RENDERING =====
function renderTable() {
    const tbody = document.getElementById('tableBody');
    const startIndex = (currentPage - 1) * (rowsPerPage === 'all' ? filteredData.length : rowsPerPage);
    const endIndex = rowsPerPage === 'all' ? filteredData.length : startIndex + rowsPerPage;
    const pageData = filteredData.slice(startIndex, endIndex);

    tbody.innerHTML = '';

    if (pageData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="${tableColumns.length + 1}" style="text-align: center; padding: 40px; color: #64748b;">
                    No data found. Try adjusting your filters.
                </td>
            </tr>
        `;
    } else {
        pageData.forEach((row, index) => {
            const actualIndex = startIndex + index;
            let rowHTML = '<tr>';
            
            // Add data cells dynamically
            tableColumns.forEach(column => {
                rowHTML += `<td>${row[column] || ''}</td>`;
            });
            
            // Add action buttons
            rowHTML += `
                <td style="white-space: nowrap;">
                    <button class="btn btn-primary btn-sm" onclick="openEditModal(${actualIndex})">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                        </svg>
                        Edit
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteRecord(${actualIndex})">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="3 6 5 6 21 6"></polyline>
                            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                        </svg>
                        Delete
                    </button>
                </td>
            </tr>`;
            
            tbody.innerHTML += rowHTML;
        });
    }

    updatePaginationInfo();
    renderPagination();
}

// ===== PAGINATION =====
function changeRowsPerPage() {
    const select = document.getElementById('rowsPerPage');
    rowsPerPage = select.value === 'all' ? 'all' : parseInt(select.value);
    currentPage = 1;
    renderTable();
}

function updatePaginationInfo() {
    const info = document.getElementById('paginationInfo');
    const startIndex = (currentPage - 1) * (rowsPerPage === 'all' ? filteredData.length : rowsPerPage) + 1;
    const endIndex = Math.min(
        rowsPerPage === 'all' ? filteredData.length : currentPage * rowsPerPage,
        filteredData.length
    );

    info.textContent = `Showing ${filteredData.length > 0 ? startIndex : 0} to ${endIndex} of ${filteredData.length} entries`;
}

function renderPagination() {
    const container = document.getElementById('paginationButtons');
    if (rowsPerPage === 'all') {
        container.innerHTML = '';
        return;
    }

    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    let html = '';

    // Previous button
    html += `<button ${currentPage === 1 ? 'disabled' : ''} onclick="changePage(${currentPage - 1})">
        ← Previous
    </button>`;

    // Page numbers
    const maxPagesToShow = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxPagesToShow / 2));
    let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);

    if (endPage - startPage < maxPagesToShow - 1) {
        startPage = Math.max(1, endPage - maxPagesToShow + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
        html += `<button class="${i === currentPage ? 'active' : ''}" onclick="changePage(${i})">
            ${i}
        </button>`;
    }

    // Next button
    html += `<button ${currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${currentPage + 1})">
        Next →
    </button>`;

    container.innerHTML = html;
}

function changePage(page) {
    currentPage = page;
    renderTable();
}

// ===== CRUD OPERATIONS =====
function openAddModal() {
    document.getElementById('modalTitle').textContent = 'Add New Record';
    document.getElementById('editForm').reset();
    document.getElementById('editIndex').value = '';
    document.getElementById('editModal').style.display = 'block';
}

function openEditModal(index) {
    const row = filteredData[index];
    document.getElementById('modalTitle').textContent = 'Edit Record';
    document.getElementById('editIndex').value = index;
    
    // Populate form fields dynamically
    tableColumns.forEach(column => {
        const fieldId = `edit${column}`;
        const field = document.getElementById(fieldId);
        if (field) {
            field.value = row[column] || '';
        }
    });
    
    document.getElementById('editModal').style.display = 'block';
}

function closeEditModal() {
    document.getElementById('editModal').style.display = 'none';
}

async function saveRecord() {
    const index = document.getElementById('editIndex').value;
    const isEdit = index !== '';

    const newRecord = {};
    tableColumns.forEach(column => {
        const fieldId = `edit${column}`;
        const field = document.getElementById(fieldId);
        if (field) {
            newRecord[column] = field.value;
        }
    });

    try {
        document.getElementById('saveButtonText').innerHTML = '<span class="loading"></span> Saving...';

        const oldRecord = isEdit ? filteredData[parseInt(index)] : null;

        if (isEdit) {
            // Find the actual index in mappingData
            const actualIndex = mappingData.findIndex(row => {
                return tableColumns.every(col => row[col] === oldRecord[col]);
            });
            mappingData[actualIndex] = newRecord;
        } else {
            // Add new record
            mappingData.push(newRecord);
        }

        // Save to Confluence
        await saveToConfluence();

        // Add to history
        await addToHistory(
            isEdit ? 'UPDATE' : 'CREATE',
            newRecord,
            oldRecord
        );

        closeEditModal();
        filterTable();
        showNotification(
            isEdit ? 'Record Updated' : 'Record Added',
            `The record has been ${isEdit ? 'updated' : 'added'} successfully.`,
            'success'
        );
    } catch (error) {
        console.error('Error saving record:', error);
        showNotification('Save Failed', 'Failed to save record. Please try again.', 'error');
    } finally {
        document.getElementById('saveButtonText').textContent = 'Save Record';
    }
}

async function deleteRecord(index) {
    if (!confirm('Are you sure you want to delete this record?')) {
        return;
    }

    try {
        const recordToDelete = filteredData[index];
        
        // Find and remove from main data
        const actualIndex = mappingData.findIndex(row => {
            return tableColumns.every(col => row[col] === recordToDelete[col]);
        });
        
        mappingData.splice(actualIndex, 1);

        // Save to Confluence
        await saveToConfluence();

        // Add to history
        await addToHistory('DELETE', recordToDelete, null);

        filterTable();
        showNotification('Record Deleted', 'The record has been deleted successfully.', 'success');
    } catch (error) {
        console.error('Error deleting record:', error);
        showNotification('Delete Failed', 'Failed to delete record. Please try again.', 'error');
    }
}

// ===== EXPORT FUNCTIONALITY =====
function exportData() {
    try {
        // Create a new workbook
        const worksheet = XLSX.utils.json_to_sheet(filteredData.length > 0 ? filteredData : mappingData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Mapping Data');
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
        const filename = `Mapping_Export_${timestamp}.xlsx`;
        
        // Download the file
        XLSX.writeFile(workbook, filename);
        
        showNotification('Export Successful', `Data exported to ${filename}`, 'success');
    } catch (error) {
        console.error('Error exporting data:', error);
        showNotification('Export Failed', 'Failed to export data. Please try again.', 'error');
    }
}

// ===== CONFLUENCE OPERATIONS =====
async function saveToConfluence() {
    // Convert data to Excel
    const worksheet = XLSX.utils.json_to_sheet(mappingData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Mapping');
    
    // Compress workbook
    const excelBuffer = XLSX.write(workbook, { 
        bookType: 'xlsx', 
        type: 'array',
        compression: true 
    });

    // Create FormData
    const formData = new FormData();
    const blob = new Blob([excelBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    formData.append('file', blob, CONFIG.mappingFileName);
    formData.append('comment', `Updated via Mapping Tool by ${currentUser}`);

    // Check if file exists
    const attachmentsUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    const attachmentsResponse = await fetch(attachmentsUrl, {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include'
    });

    const attachmentsData = await attachmentsResponse.json();
    const existingAttachment = attachmentsData.results.find(att => att.title === CONFIG.mappingFileName);

    let uploadUrl;
    if (existingAttachment) {
        // Update existing attachment
        uploadUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment/${existingAttachment.id}/data`;
    } else {
        // Create new attachment
        uploadUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    }

    const uploadResponse = await fetch(uploadUrl, {
        method: 'POST',
        headers: { 'X-Atlassian-Token': 'no-check' },
        body: formData,
        credentials: 'include'
    });

    if (!uploadResponse.ok) {
        throw new Error(`Failed to upload file: ${uploadResponse.statusText}`);
    }
}

async function addToHistory(action, newRecord, oldRecord) {
    const historyEntry = {
        Timestamp: new Date().toISOString(),
        User: currentUser,
        Action: action,
        RowID: Object.values(newRecord).slice(0, 3).join('-'),
        OldValue: oldRecord ? JSON.stringify(oldRecord) : '',
        NewValue: JSON.stringify(newRecord)
    };

    historyData.unshift(historyEntry); // Add to beginning

    // Save history to Confluence
    const worksheet = XLSX.utils.json_to_sheet(historyData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'History');
    
    const excelBuffer = XLSX.write(workbook, { 
        bookType: 'xlsx', 
        type: 'array',
        compression: true 
    });

    const formData = new FormData();
    const blob = new Blob([excelBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    formData.append('file', blob, CONFIG.historyFileName);
    formData.append('comment', `History updated by ${currentUser}`);

    // Check if history file exists
    const attachmentsUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    const attachmentsResponse = await fetch(attachmentsUrl, {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include'
    });

    const attachmentsData = await attachmentsResponse.json();
    const existingAttachment = attachmentsData.results.find(att => att.title === CONFIG.historyFileName);

    let uploadUrl;
    if (existingAttachment) {
        uploadUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment/${existingAttachment.id}/data`;
    } else {
        uploadUrl = `${CONFIG.confluenceBaseUrl}/rest/api/content/${CONFIG.pageId}/child/attachment`;
    }

    await fetch(uploadUrl, {
        method: 'POST',
        headers: { 'X-Atlassian-Token': 'no-check' },
        body: formData,
        credentials: 'include'
    });
}

// ===== HISTORY MODAL =====
function openHistoryModal() {
    const tbody = document.getElementById('historyTableBody');
    
    if (historyData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="4" style="text-align: center; padding: 40px; color: #64748b;">
                    No history records found.
                </td>
            </tr>
        `;
    } else {
        tbody.innerHTML = historyData.slice(0, 100).map(entry => {
            const timestamp = new Date(entry.Timestamp).toLocaleString();
            return `
                <tr>
                    <td>${timestamp}</td>
                    <td>${entry.User}</td>
                    <td>
                        <span class="badge ${
                            entry.Action === 'CREATE' ? 'badge-success' : 
                            entry.Action === 'UPDATE' ? 'badge-primary' : 'badge-error'
                        }">
                            ${entry.Action}
                        </span>
                    </td>
                    <td style="font-size: 12px;">${entry.RowID}</td>
                </tr>
            `;
        }).join('');
    }

    document.getElementById('historyModal').style.display = 'block';
}

function closeHistoryModal() {
    document.getElementById('historyModal').style.display = 'none';
}

// ===== STRING GENERATOR =====
function setupGeneratorListeners() {
    const generateForRadios = document.querySelectorAll('input[name="generateFor"]');
    generateForRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            const ownerSelectGroup = document.getElementById('ownerSelectGroup');
            if (this.value === 'single') {
                ownerSelectGroup.style.display = 'block';
                updateOwnersList();
            } else {
                ownerSelectGroup.style.display = 'none';
            }
        });
    });
}

function updateOwnersList() {
    const levelSelect = document.getElementById('levelSelect');
    const specificOwnerSelect = document.getElementById('specificOwnerSelect');
    const selectedLevel = levelSelect.value;

    // Get unique owners for selected level
    const owners = [...new Set(mappingData.map(row => row[selectedLevel]))].filter(o => o);
    
    specificOwnerSelect.innerHTML = '<option value="">-- Select Owner --</option>';
    owners.forEach(owner => {
        specificOwnerSelect.innerHTML += `<option value="${owner}">${owner}</option>`;
    });
}

function generateString() {
    const branchKeyField = document.getElementById('branchKeyField').value || '[branch_key]';
    const selectedLevel = document.getElementById('levelSelect').value;
    const instanceFilter = document.querySelector('input[name="instanceFilter"]:checked').value;
    const generateFor = document.querySelector('input[name="generateFor"]:checked').value;
    const specificOwner = document.getElementById('specificOwnerSelect').value;

    const instanceColumn = tableColumns.find(col => 
        col === 'TLM_INSTANCE' || col === 'INSTANCE'
    ) || 'TLM_INSTANCE';

    // Filter data based on instance
    let dataToProcess = mappingData;
    if (instanceFilter !== 'Both') {
        dataToProcess = mappingData.filter(row => row[instanceColumn] === instanceFilter);
    }

    // Group by owner
    const ownerGroups = {};
    dataToProcess.forEach(row => {
        const owner = row[selectedLevel];
        if (!owner) return;

        // If generating for single owner, skip others
        if (generateFor === 'single' && owner !== specificOwner) return;

        if (!ownerGroups[owner]) {
            ownerGroups[owner] = [];
        }

        // Find the category and branch code columns
        const categoryCol = tableColumns.find(col => col.includes('CATEGORY')) || 'CATEGORY_CODE';
        const branchCol = tableColumns.find(col => col.includes('BRANCH')) || 'BRANCH_CODE';

        // Combine CATEGORY_CODE and BRANCH_CODE
        const branchKey = `${row[categoryCol] || ''}${row[branchCol] || ''}`;
        if (branchKey) {
            ownerGroups[owner].push(branchKey);
        }
    });

    // Generate Tableau string
    let tableauString = '';
    const owners = Object.keys(ownerGroups);

    if (owners.length === 0) {
        showNotification('No Data', 'No data found for the selected criteria.', 'warning');
        return;
    }

    owners.forEach((owner, index) => {
        const branches = ownerGroups[owner];
        
        if (index > 0) {
            tableauString += 'ELSEIF ';
        } else {
            tableauString += 'IF ';
        }

        // Add conditions
        branches.forEach((branch, branchIndex) => {
            if (branchIndex > 0) {
                tableauString += ' OR\n   ';
            }
            tableauString += `${branchKeyField}="${branch}"`;
        });

        tableauString += `\nTHEN "${owner}"\n\n`;
    });

    tableauString += 'ELSE "Unknown"\nEND';

    // Display output
    document.getElementById('generatedOutput').textContent = tableauString;
    document.getElementById('outputContainer').style.display = 'block';
    
    // Scroll to output
    document.getElementById('outputContainer').scrollIntoView({ behavior: 'smooth' });
    
    showNotification('String Generated', 'Tableau calculated field generated successfully.', 'success');
}

function copyToClipboard() {
    const output = document.getElementById('generatedOutput').textContent;
    navigator.clipboard.writeText(output).then(() => {
        const btn = document.querySelector('.copy-btn');
        const originalHTML = btn.innerHTML;
        btn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="20 6 9 17 4 12"></polyline>
            </svg>
            Copied!
        `;
        showNotification('Copied', 'Text copied to clipboard successfully.', 'success');
        setTimeout(() => {
            btn.innerHTML = originalHTML;
        }, 2000);
    }).catch(err => {
        console.error('Failed to copy:', err);
        showNotification('Copy Failed', 'Failed to copy to clipboard.', 'error');
    });
}

// ===== UI HELPERS =====
function showLoadingState() {
    document.getElementById('tableBody').innerHTML = `
        <tr>
            <td colspan="${tableColumns.length + 1}" style="text-align: center; padding: 40px;">
                <div class="loading" style="width: 40px; height: 40px; margin: 0 auto;"></div>
                <p style="margin-top: 15px; color: #64748b;">Loading data...</p>
            </td>
        </tr>
    `;
}

// Close modals when clicking outside
window.onclick = function(event) {
    const editModal = document.getElementById('editModal');
    const historyModal = document.getElementById('historyModal');
    
    if (event.target === editModal) {
        closeEditModal();
    }
    if (event.target === historyModal) {
        closeHistoryModal();
    }
}
