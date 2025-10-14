// Global variables
let tableData = [];
let currentEditId = null;
let currentUser = null;

// Table column configuration
const tableColumns = [
    { key: 'tlm_instance', label: 'TLM_INSTANCE', width: '10%' },
    { key: 'categoryCode', label: 'Category Code', width: '11%' },
    { key: 'branchCode', label: 'Branch Code', width: '10%' },
    { key: 'l1Owner', label: 'L1 Owner', width: '8%' },
    { key: 'l2Owner', label: 'L2 Owner', width: '8%' },
    { key: 'l3Owner', label: 'L3 Owner', width: '8%' },
    { key: 'l4Owner', label: 'L4 Owner', width: '8%' },
    { key: 'l5Owner', label: 'L5 Owner', width: '8%' },
    { key: 'l6Owner', label: 'L6 Owner', width: '8%' },
    { key: 'l7Owner', label: 'L7 Owner', width: '8%' },
    { key: 'teamGroup', label: 'Team Group', width: '10%' },
    { key: 'actions', label: 'Actions', width: '13%' }
];

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    fetchCurrentUser();
    fetchTableData();
    setupEventListeners();
});

// Fetch current user
async function fetchCurrentUser() {
    try {
        // Replace with your actual API endpoint
        const response = await fetch('/api/current-user', {
            headers: {
                'Authorization': 'Bearer ' + localStorage.getItem('token')
            }
        });
        
        if (response.ok) {
            const userData = await response.json();
            currentUser = userData;
            document.getElementById('username').textContent = userData.username || userData.name || 'User';
        } else {
            // Fallback for demo
            document.getElementById('username').textContent = 'Demo User';
        }
    } catch (error) {
        console.error('Error fetching user:', error);
        // Demo fallback
        document.getElementById('username').textContent = 'Demo User';
    }
}

// Fetch table data
async function fetchTableData() {
    showNotification('Loading data...', 'info');
    
    try {
        // Replace with your actual API endpoint
        const response = await fetch('/api/tlm-instances');
        
        if (response.ok) {
            tableData = await response.json();
        } else {
            // Demo data fallback
            tableData = getDemoData();
        }
    } catch (error) {
        console.error('Error fetching data:', error);
        // Use demo data as fallback
        tableData = getDemoData();
    }
    
    renderTable();
    showNotification('Data loaded successfully', 'success');
}

// Demo data
function getDemoData() {
    return [
        {
            id: 1,
            tlm_instance: 'CASH',
            categoryCode: 'IT.CASH',
            branchCode: 'CODE_NR1',
            l1Owner: 'Adam',
            l2Owner: 'Gira',
            l3Owner: 'Ahmadjon',
            l4Owner: 'Eoin',
            l5Owner: 'Dosia',
            l6Owner: 'Hugo',
            l7Owner: 'Karishma',
            teamGroup: 'LATAM INCOME'
        },
        {
            id: 2,
            tlm_instance: 'CASH',
            categoryCode: 'DE.CASH',
            branchCode: 'CODE_NR1',
            l1Owner: 'Adam',
            l2Owner: 'Gira',
            l3Owner: 'Ahmadjon',
            l4Owner: 'Eoin',
            l5Owner: 'Dosia',
            l6Owner: 'Hugo',
            l7Owner: 'Karishma',
            teamGroup: 'LATAM INCOME'
        },
        {
            id: 3,
            tlm_instance: 'STOCK',
            categoryCode: 'BR.ST',
            branchCode: 'CODE_NR11',
            l1Owner: 'David',
            l2Owner: 'Tiro',
            l3Owner: 'Bob',
            l4Owner: 'Eoin',
            l5Owner: 'Dosia',
            l6Owner: 'Hugo',
            l7Owner: 'Karishma',
            teamGroup: 'TRADING TEAM'
        },
        {
            id: 4,
            tlm_instance: 'STOCK',
            categoryCode: 'US.ST',
            branchCode: 'CODE_NR24',
            l1Owner: 'David',
            l2Owner: 'Tiro',
            l3Owner: 'Bob',
            l4Owner: 'Eoin',
            l5Owner: 'Dosia',
            l6Owner: 'Hugo',
            l7Owner: 'Karishma',
            teamGroup: 'TRADING TEAM'
        },
        {
            id: 5,
            tlm_instance: 'CASH',
            categoryCode: 'SE.CASH',
            branchCode: 'CODE_NR2',
            l1Owner: 'Cecil',
            l2Owner: 'Dominik',
            l3Owner: 'Ahmadjon',
            l4Owner: 'Eoin',
            l5Owner: 'Dosia',
            l6Owner: 'Hugo',
            l7Owner: 'Karishma',
            teamGroup: 'TAX TEAM'
        }
    ];
}

// Render table
function renderTable() {
    const table = document.getElementById('dataTable');
    table.innerHTML = '';
    
    // Create thead
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    tableColumns.forEach(column => {
        const th = document.createElement('th');
        th.style.width = column.width;
        th.textContent = column.label;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create tbody
    const tbody = document.createElement('tbody');
    
    tableData.forEach(row => {
        const tr = document.createElement('tr');
        
        tableColumns.forEach(column => {
            const td = document.createElement('td');
            td.style.width = column.width;
            
            if (column.key === 'actions') {
                // Create action buttons
                const actionsDiv = document.createElement('div');
                actionsDiv.className = 'actions-cell';
                
                const editBtn = document.createElement('button');
                editBtn.className = 'btn btn-sm btn-edit';
                editBtn.textContent = '✏️ Edit';
                editBtn.onclick = () => handleEdit(row.id);
                
                const deleteBtn = document.createElement('button');
                deleteBtn.className = 'btn btn-sm btn-delete';
                deleteBtn.textContent = '🗑️ Delete';
                deleteBtn.onclick = () => handleDelete(row.id);
                
                actionsDiv.appendChild(editBtn);
                actionsDiv.appendChild(deleteBtn);
                td.appendChild(actionsDiv);
            } else if (column.key === 'tlm_instance') {
                // Create badge for TLM_INSTANCE
                const badge = document.createElement('span');
                badge.className = `instance-badge ${row[column.key].toLowerCase()}`;
                badge.textContent = row[column.key];
                td.appendChild(badge);
            } else {
                td.textContent = row[column.key] || '';
            }
            
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });
    
    table.appendChild(tbody);
}

// Setup event listeners
function setupEventListeners() {
    // Add new button
    document.getElementById('addNewBtn').addEventListener('click', handleAddNew);
    
    // Modal close button
    document.getElementById('closeModal').addEventListener('click', closeModal);
    document.getElementById('cancelEdit').addEventListener('click', closeModal);
    
    // Form submit
    document.getElementById('editForm').addEventListener('submit', handleFormSubmit);
    
    // Close modal on outside click
    window.addEventListener('click', function(event) {
        const modal = document.getElementById('editModal');
        if (event.target === modal) {
            closeModal();
        }
    });
}

// Handle add new
function handleAddNew() {
    currentEditId = null;
    openModal();
    clearForm();
    showNotification('Adding new record', 'info');
}

// Handle edit
function handleEdit(id) {
    currentEditId = id;
    const record = tableData.find(item => item.id === id);
    
    if (record) {
        openModal();
        populateForm(record);
        showNotification('Editing record', 'info');
    }
}

// Handle delete
async function handleDelete(id) {
    if (confirm('Are you sure you want to delete this record?')) {
        showNotification('Deleting record...', 'warning');
        
        try {
            // Replace with actual API call
            const response = await fetch(`/api/tlm-instances/${id}`, {
                method: 'DELETE',
                headers: {
                    'Authorization': 'Bearer ' + localStorage.getItem('token')
                }
            });
            
            if (response.ok || true) { // true for demo
                // Remove from local data
                tableData = tableData.filter(item => item.id !== id);
                renderTable();
                showNotification('Record deleted successfully', 'success');
            } else {
                showNotification('Failed to delete record', 'error');
            }
        } catch (error) {
            console.error('Error deleting record:', error);
            // Demo mode - still delete locally
            tableData = tableData.filter(item => item.id !== id);
            renderTable();
            showNotification('Record deleted successfully', 'success');
        }
    }
}

// Handle form submit
async function handleFormSubmit(e) {
    e.preventDefault();
    
    const formData = {
        tlm_instance: document.getElementById('editInstance').value,
        categoryCode: document.getElementById('editCategoryCode').value,
        branchCode: document.getElementById('editBranchCode').value,
        l1Owner: document.getElementById('editL1Owner').value,
        l2Owner: document.getElementById('editL2Owner').value,
        l3Owner: document.getElementById('editL3Owner').value,
        l4Owner: document.getElementById('editL4Owner').value,
        l5Owner: document.getElementById('editL5Owner').value,
        l6Owner: document.getElementById('editL6Owner').value,
        l7Owner: document.getElementById('editL7Owner').value,
        teamGroup: document.getElementById('editTeamGroup').value
    };
    
    showNotification('Saving changes...', 'info');
    
    try {
        if (currentEditId) {
            // Update existing record
            const response = await fetch(`/api/tlm-instances/${currentEditId}`, {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + localStorage.getItem('token')
                },
                body: JSON.stringify(formData)
            });
            
            if (response.ok || true) { // true for demo
                // Update local data
                const index = tableData.findIndex(item => item.id === currentEditId);
                if (index !== -1) {
                    tableData[index] = { ...tableData[index], ...formData };
                }
                showNotification('Record updated successfully', 'success');
            } else {
                showNotification('Failed to update record', 'error');
            }
        } else {
            // Add new record
            const response = await fetch('/api/tlm-instances', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + localStorage.getItem('token')
                },
                body: JSON.stringify(formData)
            });
            
            if (response.ok || true) { // true for demo
                // Add to local data
                const newId = Math.max(...tableData.map(item => item.id), 0) + 1;
                tableData.push({ id: newId, ...formData });
                showNotification('Record added successfully', 'success');
            } else {
                showNotification('Failed to add record', 'error');
            }
        }
        
        renderTable();
        closeModal();
    } catch (error) {
        console.error('Error saving record:', error);
        // Demo mode - still save locally
        if (currentEditId) {
            const index = tableData.findIndex(item => item.id === currentEditId);
            if (index !== -1) {
                tableData[index] = { ...tableData[index], ...formData };
            }
            showNotification('Record updated successfully', 'success');
        } else {
            const newId = Math.max(...tableData.map(item => item.id), 0) + 1;
            tableData.push({ id: newId, ...formData });
            showNotification('Record added successfully', 'success');
        }
        renderTable();
        closeModal();
    }
}

// Modal functions
function openModal() {
    document.getElementById('editModal').classList.add('show');
}

function closeModal() {
    document.getElementById('editModal').classList.remove('show');
    clearForm();
}

function clearForm() {
    document.getElementById('editForm').reset();
}

function populateForm(record) {
    document.getElementById('editInstance').value = record.tlm_instance;
    document.getElementById('editCategoryCode').value = record.categoryCode;
    document.getElementById('editBranchCode').value = record.branchCode;
    document.getElementById('editL1Owner').value = record.l1Owner;
    document.getElementById('editL2Owner').value = record.l2Owner;
    document.getElementById('editL3Owner').value = record.l3Owner;
    document.getElementById('editL4Owner').value = record.l4Owner;
    document.getElementById('editL5Owner').value = record.l5Owner;
    document.getElementById('editL6Owner').value = record.l6Owner;
    document.getElementById('editL7Owner').value = record.l7Owner;
    document.getElementById('editTeamGroup').value = record.teamGroup;
}

// Notification system
function showNotification(message, type = 'info', duration = 3000) {
    const container = document.getElementById('notification-container');
    
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    
    // Icon based on type
    const icons = {
        success: '✓',
        error: '✕',
        warning: '⚠',
        info: 'ℹ'
    };
    
    notification.innerHTML = `
        <span class="notification-icon">${icons[type]}</span>
        <span class="notification-message">${message}</span>
        <span class="notification-close">&times;</span>
    `;
    
    // Add to container
    container.appendChild(notification);
    
    // Close button functionality
    notification.querySelector('.notification-close').addEventListener('click', function() {
        removeNotification(notification);
    });
    
    // Auto remove after duration
    setTimeout(() => {
        removeNotification(notification);
    }, duration);
}

function removeNotification(notification) {
    notification.classList.add('hiding');
    setTimeout(() => {
        notification.remove();
    }, 300);
}

// Export for debugging
window.debugData = {
    tableData,
    currentUser,
    showNotification
};