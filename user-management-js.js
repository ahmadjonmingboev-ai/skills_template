// User Management Page JavaScript

// Configuration
const CONFIG = {
    confluence: {
        domain: 'https://www.myconfluence.net',  // Update with actual domain
        pageId: '123345346',
        api: {
            currentUser: '/rest/api/user/current',
            attachment: '/rest/api/content/{pageId}/child/attachment'
        }
    },
    files: {
        userDatabase: 'userdatabase.xlsx',
        mainTaskDatabase: 'maintaskdatabase.xlsx',
        approvals: 'approvals.xlsx',
        auditLog: 'audit_log.xlsx'
    },
    maxRowsPerPage: 50
};

// API Helper Functions
const API = {
    getCurrentUser: async function() {
        try {
            const response = await fetch(CONFIG.confluence.domain + CONFIG.confluence.api.currentUser, {
                credentials: 'same-origin'
            });
            if (!response.ok) {
                console.warn('User API returned:', response.status);
                return null;
            }
            const userData = await response.json();
            return {
                displayName: userData.displayName || userData.fullName || userData.name || 'Unknown',
                key: userData.key || userData.username || userData.accountId || userData.name || userData.userName,
                email: userData.emailAddress || userData.email,
                ...userData
            };
        } catch (error) {
            console.warn('Error fetching current user:', error);
            return null;
        }
    },
    
    getAttachments: async function() {
        try {
            const url = CONFIG.confluence.domain +
                       CONFIG.confluence.api.attachment.replace('{pageId}', CONFIG.confluence.pageId);
            const response = await fetch(url, {
                credentials: 'same-origin'
            });
            if (!response.ok) throw new Error('Failed to fetch attachments');
            const data = await response.json();
            return data.results || [];
        } catch (error) {
            console.error('Error fetching attachments:', error);
            return [];
        }
    },
    
    downloadAttachment: async function(attachmentUrl) {
        try {
            const url = attachmentUrl.startsWith('http') ? attachmentUrl : CONFIG.confluence.domain + attachmentUrl;
            const response = await fetch(url, {
                credentials: 'same-origin',
                headers: {
                    'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                }
            });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            return await response.arrayBuffer();
        } catch (error) {
            console.error('Error downloading attachment:', error);
            return null;
        }
    },
    
    uploadAttachment: async function(filename, blob) {
        try {
            const attachments = await this.getAttachments();
            const existingAttachment = attachments.find(att => att.title === filename);
            
            if (!existingAttachment) {
                throw new Error(`Attachment ${filename} not found`);
            }
            
            const formData = new FormData();
            formData.append('file', blob, filename);
            formData.append('minorEdit', 'true');
            formData.append('comment', 'Updated by User Management System');
            
            let url, response;
            
            // Try updating existing attachment
            try {
                url = `${CONFIG.confluence.domain}/rest/api/content/${existingAttachment.id}/child/attachment/${existingAttachment.id}/data`;
                response = await fetch(url, {
                    method: 'POST',
                    body: formData,
                    credentials: 'same-origin',
                    headers: {
                        'X-Atlassian-Token': 'no-check'
                    }
                });
                
                if (response.ok) return await response.json();
            } catch (e) {
                console.log('Method 1 failed, trying method 2');
            }
            
            // Try creating new version
            try {
                url = `${CONFIG.confluence.domain}/rest/api/content/${CONFIG.confluence.pageId}/child/attachment`;
                response = await fetch(url, {
                    method: 'POST',
                    body: formData,
                    credentials: 'same-origin',
                    headers: {
                        'X-Atlassian-Token': 'no-check'
                    }
                });
                
                if (response.ok) return await response.json();
            } catch (e) {
                console.log('Method 2 failed');
            }
            
            throw new Error('Failed to upload attachment');
            
        } catch (error) {
            console.error('Error uploading attachment:', error);
            throw error;
        }
    }
};

// Global variables
let currentUser = null;
let allUsers = [];
let filteredUsers = [];
let currentPage = 1;
let attachments = {};
let userToDelete = null;
let mainWorkbook = null;
let templateData = [];

// Initialize the application
async function initializeApp() {
    showLoading(true, 'Loading user management...');
    try {
        // Get current user from Confluence
        const confluenceUser = await API.getCurrentUser();
        if (!confluenceUser) {
            showToast('Could not get current user. Please ensure you are logged into Confluence.', 'error');
            return;
        }
        
        currentUser = {
            username: confluenceUser.key,
            displayName: confluenceUser.displayName,
            email: confluenceUser.email
        };
        
        // Load all attachments first
        await loadAttachments();
        
        // Check if user is admin
        const hasAccess = await checkAdminAccess();
        if (!hasAccess) {
            showNoAccess();
            return;
        }
        
        // Load main workbook for template data
        await loadMainWorkbook();
        
        // Load user data
        await loadUserData();
        
        // Initialize UI
        initializeUI();
        
    } catch (error) {
        console.error('Initialization error:', error);
        showToast('Failed to initialize application. Please refresh the page.', 'error');
    } finally {
        showLoading(false);
    }
}

// Load all attachments
async function loadAttachments() {
    const attachmentList = await API.getAttachments();
    attachmentList.forEach(att => {
        attachments[att.title] = att._links.download || att._links.webui;
    });
}

// Fetch and parse Excel file
async function fetchExcelData(filename) {
    try {
        const downloadUrl = attachments[filename];
        if (!downloadUrl) {
            throw new Error(`Attachment ${filename} not found`);
        }
        
        const arrayBuffer = await API.downloadAttachment(downloadUrl);
        if (!arrayBuffer) {
            throw new Error(`Failed to download ${filename}`);
        }
        
        return XLSX.read(arrayBuffer, { type: 'array' });
        
    } catch (error) {
        console.error(`Error fetching ${filename}:`, error);
        throw error;
    }
}

// Check admin access
async function checkAdminAccess() {
    try {
        const workbook = await fetchExcelData(CONFIG.files.userDatabase);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const users = XLSX.utils.sheet_to_json(worksheet);
        
        const userData = users.find(u => u.username === currentUser.username);
        if (!userData) {
            // Fallback for testing - remove in production
            return users.find(u => u.username === 'BD12345')?.role === 'admin';
        }
        
        currentUser = { ...currentUser, ...userData };
        return userData.role === 'admin';
        
    } catch (error) {
        console.error('Failed to check admin access:', error);
        return false;
    }
}

// Load main workbook for template
async function loadMainWorkbook() {
    try {
        mainWorkbook = await fetchExcelData(CONFIG.files.mainTaskDatabase);
        const templateSheet = mainWorkbook.Sheets['Template'];
        if (templateSheet) {
            templateData = XLSX.utils.sheet_to_json(templateSheet);
        }
        console.log('Template loaded with', templateData.length, 'tasks');
    } catch (error) {
        console.error('Failed to load main workbook:', error);
    }
}

// Load user data
async function loadUserData() {
    try {
        const workbook = await fetchExcelData(CONFIG.files.userDatabase);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        allUsers = XLSX.utils.sheet_to_json(worksheet);
        
        // Update statistics
        const teams = [...new Set(allUsers.map(u => u.team))].filter(Boolean);
        document.getElementById('totalUsers').textContent = allUsers.length;
        document.getElementById('totalTeams').textContent = teams.length;
        
        console.log('Loaded users:', allUsers.length);
        
    } catch (error) {
        console.error('Failed to load user data:', error);
        allUsers = [];
    }
}

// Initialize UI
function initializeUI() {
    setupFilters();
    filterUsers();
}

// Setup filter dropdowns
function setupFilters() {
    const teams = [...new Set(allUsers.map(u => u.team))].filter(Boolean);
    
    const teamFilter = document.getElementById('teamFilter');
    teamFilter.innerHTML = '<option value="">All Teams</option>';
    teams.forEach(team => {
        const option = document.createElement('option');
        option.value = team;
        option.textContent = team;
        teamFilter.appendChild(option);
    });
}

// Filter users
function filterUsers() {
    const searchTerm = document.getElementById('userSearchInput').value.toLowerCase();
    const teamFilter = document.getElementById('teamFilter').value;
    const roleFilter = document.getElementById('roleFilter').value;
    
    filteredUsers = allUsers.filter(user => {
        const matchesSearch = !searchTerm || 
            (user.username && user.username.toLowerCase().includes(searchTerm)) ||
            (user.name && user.name.toLowerCase().includes(searchTerm)) ||
            (user.email && user.email.toLowerCase().includes(searchTerm));
        
        const matchesTeam = !teamFilter || user.team === teamFilter;
        const matchesRole = !roleFilter || user.role === roleFilter;
        
        return matchesSearch && matchesTeam && matchesRole;
    });
    
    currentPage = 1;
    renderUsersTable();
}

// Render users table
function renderUsersTable() {
    const tbody = document.getElementById('usersTableBody');
    const emptyState = document.getElementById('usersEmptyState');
    
    tbody.innerHTML = '';
    
    if (filteredUsers.length === 0) {
        document.getElementById('usersTable').style.display = 'none';
        emptyState.style.display = 'flex';
        updatePaginationInfo(0, 0, 0);
        return;
    }
    
    document.getElementById('usersTable').style.display = 'table';
    emptyState.style.display = 'none';
    
    // Paginate
    const startIndex = (currentPage - 1) * CONFIG.maxRowsPerPage;
    const endIndex = Math.min(startIndex + CONFIG.maxRowsPerPage, filteredUsers.length);
    const pageData = filteredUsers.slice(startIndex, endIndex);
    
    // Render rows
    pageData.forEach((user, index) => {
        const row = document.createElement('tr');
        
        const roleColor = {
            admin: 'role-admin',
            manager: 'role-manager',
            user: 'role-user'
        };
        
        row.innerHTML = `
            <td>${user.username || ''}</td>
            <td>${user.name || ''}</td>
            <td>${user.email || ''}</td>
            <td>${user.team || ''}</td>
            <td><span class="role-badge ${roleColor[user.role] || ''}">${user.role || ''}</span></td>
            <td>${user.tenure || ''}</td>
            <td>
                <div class="user-actions">
                    <button class="user-action-btn edit" onclick="editUserByIndex(${startIndex + index})">
                        <svg fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/>
                        </svg>
                        Edit
                    </button>
                    <button class="user-action-btn delete" onclick="deleteUserByIndex(${startIndex + index})">
                        <svg fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/>
                        </svg>
                        Delete
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });
    
    updatePagination();
    updatePaginationInfo(startIndex + 1, endIndex, filteredUsers.length);
}

// Edit user by index
function editUserByIndex(index) {
    const user = filteredUsers[index];
    if (user) {
        editUser(user);
    }
}

// Delete user by index
function deleteUserByIndex(index) {
    const user = filteredUsers[index];
    if (user) {
        deleteUser(user);
    }
}

// Open add user modal
function openAddUserModal() {
    document.getElementById('userModalTitle').textContent = 'Add New User';
    document.getElementById('userForm').reset();
    document.getElementById('userEditMode').value = 'false';
    document.getElementById('originalUsername').value = '';
    document.getElementById('userModal').classList.add('active');
}

// Edit user
function editUser(user) {
    document.getElementById('userModalTitle').textContent = 'Edit User';
    document.getElementById('userEditMode').value = 'true';
    document.getElementById('originalUsername').value = user.username;
    
    document.getElementById('userUsername').value = user.username || '';
    document.getElementById('userName').value = user.name || '';
    document.getElementById('userEmail').value = user.email || '';
    document.getElementById('userTeam').value = user.team || '';
    document.getElementById('userRole').value = user.role || '';
    document.getElementById('userTenure').value = user.tenure || '';
    
    document.getElementById('userModal').classList.add('active');
}

// Save user
async function saveUser() {
    const isEditMode = document.getElementById('userEditMode').value === 'true';
    const originalUsername = document.getElementById('originalUsername').value;
    
    const userData = {
        username: document.getElementById('userUsername').value.trim(),
        name: document.getElementById('userName').value.trim(),
        email: document.getElementById('userEmail').value.trim(),
        team: document.getElementById('userTeam').value.trim(),
        role: document.getElementById('userRole').value,
        tenure: document.getElementById('userTenure').value.trim()
    };
    
    // Validation
    if (!userData.username || !userData.name || !userData.email || !userData.team || !userData.role) {
        showToast('Please fill in all required fields', 'warning');
        return;
    }
    
    // Check for duplicate username (only for new users or if username changed)
    if (!isEditMode || userData.username !== originalUsername) {
        const existingUser = allUsers.find(u => u.username === userData.username);
        if (existingUser) {
            showToast('Username already exists. Please use a different username.', 'error');
            return;
        }
    }
    
    showLoading(true, isEditMode ? 'Updating user...' : 'Adding user...');
    
    try {
        if (isEditMode) {
            await updateExistingUser(originalUsername, userData);
        } else {
            await addNewUser(userData);
        }
        
        // Log to audit
        await logAudit(
            isEditMode ? 'EDIT_USER' : 'ADD_USER',
            userData.username,
            isEditMode ? `Updated user: ${userData.name}` : `Added new user: ${userData.name}`,
            isEditMode ? originalUsername : '',
            JSON.stringify(userData)
        );
        
        closeUserModal();
        showToast(`User ${isEditMode ? 'updated' : 'added'} successfully`, 'success');
        
        // Reload data
        await loadUserData();
        initializeUI();
        
    } catch (error) {
        console.error('Failed to save user:', error);
        showToast(`Failed to ${isEditMode ? 'update' : 'add'} user`, 'error');
    } finally {
        showLoading(false);
    }
}

// Add new user
async function addNewUser(userData) {
    // Update user database
    const userWorkbook = await fetchExcelData(CONFIG.files.userDatabase);
    const userSheet = userWorkbook.Sheets[userWorkbook.SheetNames[0]];
    const users = XLSX.utils.sheet_to_json(userSheet);
    users.push(userData);
    userWorkbook.Sheets[userWorkbook.SheetNames[0]] = XLSX.utils.json_to_sheet(users);
    
    const userOut = XLSX.write(userWorkbook, { bookType: 'xlsx', type: 'array' });
    const userBlob = new Blob([userOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await API.uploadAttachment(CONFIG.files.userDatabase, userBlob);
    
    // Add user sheet to main task database
    updateLoadingMessage('Creating user skill sheet...');
    
    // Create new sheet for user with template data
    const newSheetData = templateData.map(task => {
        const userTask = { ...task };
        // Get all market columns (those not in standard columns)
        const standardColumns = ['Item_ID', 'Category', 'Task_Group', 'Task_Name'];
        Object.keys(task).forEach(key => {
            if (!standardColumns.includes(key)) {
                userTask[key] = 'no';  // Set all markets to 'no' by default
            }
        });
        return userTask;
    });
    
    // Add new sheet to workbook
    const newSheet = XLSX.utils.json_to_sheet(newSheetData);
    mainWorkbook.Sheets[userData.username] = newSheet;
    
    // Update SheetNames array if not present
    if (!mainWorkbook.SheetNames.includes(userData.username)) {
        mainWorkbook.SheetNames.push(userData.username);
    }
    
    // Save main workbook
    const mainOut = XLSX.write(mainWorkbook, { bookType: 'xlsx', type: 'array' });
    const mainBlob = new Blob([mainOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await API.uploadAttachment(CONFIG.files.mainTaskDatabase, mainBlob);
}

// Update existing user
async function updateExistingUser(originalUsername, userData) {
    // Update user database
    const userWorkbook = await fetchExcelData(CONFIG.files.userDatabase);
    const userSheet = userWorkbook.Sheets[userWorkbook.SheetNames[0]];
    const users = XLSX.utils.sheet_to_json(userSheet);
    
    const userIndex = users.findIndex(u => u.username === originalUsername);
    if (userIndex !== -1) {
        users[userIndex] = userData;
    }
    
    userWorkbook.Sheets[userWorkbook.SheetNames[0]] = XLSX.utils.json_to_sheet(users);
    
    const userOut = XLSX.write(userWorkbook, { bookType: 'xlsx', type: 'array' });
    const userBlob = new Blob([userOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await API.uploadAttachment(CONFIG.files.userDatabase, userBlob);
    
    // If username changed, rename sheet in main workbook
    if (originalUsername !== userData.username) {
        updateLoadingMessage('Renaming user skill sheet...');
        
        // Get the sheet data
        const oldSheet = mainWorkbook.Sheets[originalUsername];
        if (oldSheet) {
            // Copy sheet to new name
            mainWorkbook.Sheets[userData.username] = oldSheet;
            
            // Delete old sheet
            delete mainWorkbook.Sheets[originalUsername];
            
            // Update SheetNames array
            const sheetIndex = mainWorkbook.SheetNames.indexOf(originalUsername);
            if (sheetIndex !== -1) {
                mainWorkbook.SheetNames[sheetIndex] = userData.username;
            }
            
            // Save main workbook
            const mainOut = XLSX.write(mainWorkbook, { bookType: 'xlsx', type: 'array' });
            const mainBlob = new Blob([mainOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            await API.uploadAttachment(CONFIG.files.mainTaskDatabase, mainBlob);
        }
    }
}

// Delete user
function deleteUser(user) {
    userToDelete = user;
    document.getElementById('deleteUserName').textContent = user.name;
    document.getElementById('deleteUsername').textContent = user.username;
    document.getElementById('deleteUserModal').classList.add('active');
}

// Confirm delete user
async function confirmDeleteUser() {
    if (!userToDelete) return;
    
    showLoading(true, 'Deleting user...');
    
    try {
        // Remove from user database
        updateLoadingMessage('Removing from user database...');
        const userWorkbook = await fetchExcelData(CONFIG.files.userDatabase);
        const userSheet = userWorkbook.Sheets[userWorkbook.SheetNames[0]];
        let users = XLSX.utils.sheet_to_json(userSheet);
        users = users.filter(u => u.username !== userToDelete.username);
        userWorkbook.Sheets[userWorkbook.SheetNames[0]] = XLSX.utils.json_to_sheet(users);
        
        const userOut = XLSX.write(userWorkbook, { bookType: 'xlsx', type: 'array' });
        const userBlob = new Blob([userOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        await API.uploadAttachment(CONFIG.files.userDatabase, userBlob);
        
        // Remove sheet from main workbook
        updateLoadingMessage('Removing skill data...');
        if (mainWorkbook.Sheets[userToDelete.username]) {
            delete mainWorkbook.Sheets[userToDelete.username];
            
            const sheetIndex = mainWorkbook.SheetNames.indexOf(userToDelete.username);
            if (sheetIndex !== -1) {
                mainWorkbook.SheetNames.splice(sheetIndex, 1);
            }
            
            const mainOut = XLSX.write(mainWorkbook, { bookType: 'xlsx', type: 'array' });
            const mainBlob = new Blob([mainOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            await API.uploadAttachment(CONFIG.files.mainTaskDatabase, mainBlob);
        }
        
        // Remove from approvals if any
        updateLoadingMessage('Cleaning up approvals...');
        try {
            const approvalsWorkbook = await fetchExcelData(CONFIG.files.approvals);
            const approvalsSheet = approvalsWorkbook.Sheets[approvalsWorkbook.SheetNames[0]];
            let approvals = XLSX.utils.sheet_to_json(approvalsSheet);
            const originalCount = approvals.length;
            approvals = approvals.filter(a => a.username !== userToDelete.username);
            
            if (approvals.length < originalCount) {
                approvalsWorkbook.Sheets[approvalsWorkbook.SheetNames[0]] = XLSX.utils.json_to_sheet(approvals);
                const approvalsOut = XLSX.write(approvalsWorkbook, { bookType: 'xlsx', type: 'array' });
                const approvalsBlob = new Blob([approvalsOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                await API.uploadAttachment(CONFIG.files.approvals, approvalsBlob);
            }
        } catch (e) {
            console.log('No approvals to clean up');
        }
        
        // Log to audit
        await logAudit(
            'DELETE_USER',
            userToDelete.username,
            `Deleted user: ${userToDelete.name}`,
            JSON.stringify(userToDelete),
            ''
        );
        
        closeDeleteUserModal();
        showToast('User deleted successfully', 'success');
        
        // Reload data
        await loadUserData();
        initializeUI();
        
    } catch (error) {
        console.error('Failed to delete user:', error);
        showToast('Failed to delete user', 'error');
    } finally {
        showLoading(false);
    }
}

// Log to audit
async function logAudit(action, itemId, details, previousValue = '', newValue = '') {
    try {
        let auditData = [];
        
        // Try to load existing audit log
        try {
            const workbook = await fetchExcelData(CONFIG.files.auditLog);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            auditData = XLSX.utils.sheet_to_json(worksheet);
        } catch (e) {
            console.log('No existing audit log, creating new one');
        }
        
        // Add new audit entry
        auditData.push({
            Timestamp: new Date().toISOString(),
            User: currentUser.username,
            Action: action,
            Item_ID: itemId,
            Details: details,
            Previous_Value: previousValue,
            New_Value: newValue
        });
        
        // Save audit log
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(auditData);
        XLSX.utils.book_append_sheet(wb, ws, 'Audit Log');
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        await API.uploadAttachment(CONFIG.files.auditLog, blob);
        
    } catch (error) {
        console.error('Failed to log audit:', error);
    }
}

// Export users
function exportUsers() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(filteredUsers);
    XLSX.utils.book_append_sheet(wb, ws, 'Users');
    XLSX.writeFile(wb, `users_export_${new Date().getTime()}.xlsx`);
    showToast('Users exported successfully', 'success');
}

// Pagination functions
function updatePagination() {
    const totalPages = Math.ceil(filteredUsers.length / CONFIG.maxRowsPerPage);
    const controls = document.getElementById('userPaginationControls');
    controls.innerHTML = '';
    
    if (totalPages <= 1) return;
    
    // Previous button
    const prevBtn = document.createElement('button');
    prevBtn.className = 'page-btn';
    prevBtn.textContent = 'Previous';
    prevBtn.disabled = currentPage === 1;
    prevBtn.onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            renderUsersTable();
        }
    };
    controls.appendChild(prevBtn);
    
    // Page numbers
    for (let i = 1; i <= Math.min(totalPages, 5); i++) {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'page-btn' + (i === currentPage ? ' active' : '');
        pageBtn.textContent = i;
        pageBtn.onclick = () => {
            currentPage = i;
            renderUsersTable();
        };
        controls.appendChild(pageBtn);
    }
    
    // Next button
    const nextBtn = document.createElement('button');
    nextBtn.className = 'page-btn';
    nextBtn.textContent = 'Next';
    nextBtn.disabled = currentPage === totalPages;
    nextBtn.onclick = () => {
        if (currentPage < totalPages) {
            currentPage++;
            renderUsersTable();
        }
    };
    controls.appendChild(nextBtn);
}

function updatePaginationInfo(from, to, total) {
    document.getElementById('userShowingFrom').textContent = from;
    document.getElementById('userShowingTo').textContent = to;
    document.getElementById('userTotalRecords').textContent = total;
}

// Modal functions
function closeUserModal() {
    document.getElementById('userModal').classList.remove('active');
    document.getElementById('userForm').reset();
}

function closeDeleteUserModal() {
    document.getElementById('deleteUserModal').classList.remove('active');
    userToDelete = null;
}

// Refresh data
async function refreshUserData() {
    await initializeApp();
    showToast('Data refreshed successfully', 'success');
}

// Utility functions
function showLoading(show, message = 'Loading...') {
    const overlay = document.getElementById('userLoadingOverlay');
    if (overlay) {
        overlay.classList.toggle('active', show);
        if (message) {
            updateLoadingMessage(message);
        }
    }
}

function updateLoadingMessage(message) {
    const messageElement = document.getElementById('userLoadingMessage');
    if (messageElement) {
        messageElement.textContent = message;
    }
}

function showNoAccess() {
    document.getElementById('userManagementPage').style.display = 'none';
    document.getElementById('userNoAccessMessage').style.display = 'flex';
}

// Toast notification system
function showToast(message, type = 'info') {
    const toastContainer = document.getElementById('userToastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    const icons = {
        success: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>',
        error: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>',
        warning: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>',
        info: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>'
    };
    
    toast.innerHTML = `
        ${icons[type]}
        <span style="margin-left: 8px;">${message}</span>
    `;
    
    toastContainer.appendChild(toast);
    
    setTimeout(() => {
        toast.classList.add('removing');
        setTimeout(() => {
            if (toast.parentElement) {
                toast.remove();
            }
        }, 300);
    }, 3000);
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}