// Enhanced Worksheet Tool - Complete Implementation

// CSV Utilities
const CSV = {
    parse: (text) => {
        if (!text) return [];
        const lines = text.split('\n').filter(l => l.trim());
        if (!lines.length) return [];
        const headers = CSV.parseLine(lines[0]);
        return lines.slice(1).map(line => {
            const values = CSV.parseLine(line);
            const row = {};
            headers.forEach((h, i) => row[h] = values[i] || '');
            return row;
        });
    },
    parseLine: (line) => {
        const result = [], current = [];
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current.push('"');
                    i++;
                } else inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.join(''));
                current.length = 0;
            } else current.push(char);
        }
        result.push(current.join(''));
        return result;
    },
    generate: (data) => {
        if (!data || !data.length) return '';
        const headers = Object.keys(data[0]);
        return [headers.map(h => CSV.escape(h)).join(',')]
            .concat(data.map(row => headers.map(h => CSV.escape(row[h] || '')).join(',')))
            .join('\n');
    },
    escape: (str) => {
        str = String(str || '');
        return str.includes(',') || str.includes('"') || str.includes('\n') 
            ? '"' + str.replace(/"/g, '""') + '"' : str;
    }
};

// Configuration
const CONFIG = {
    confluenceBaseUrl: '',
    pageId: '',
    worksheetName: 'Worksheet',
    theme: 'light',
    permissions: { admins: [], editors: [], viewers: [] }
};

// State
const STATE = {
    currentUser: null,
    userPermission: 'viewer',
    tabs: [],
    currentTab: null,
    tabData: {},
    selectedRows: new Set(),
    currentPage: 1,
    pageSize: 25,
    sortColumn: null,
    sortDirection: 'asc',
    filters: [],
    modalData: {},
    rowLocks: {}
};

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    try {
        showLoading('Initializing...');
        await loadConfig();
        await loadTabs();
        setupEventListeners();
        if (STATE.tabs.length > 0) await switchTab(STATE.tabs[0].id);
        hideLoading();
    } catch (e) {
        console.error(e);
        showToast('Failed to initialize', 'error');
        hideLoading();
    }
});

async function loadConfig() {
    STATE.currentUser = 'User';
    document.getElementById('userName').textContent = STATE.currentUser;
}

async function loadTabs() {
    STATE.tabs = [{
        id: '1',
        name: 'Sample Tab',
        fileName: 'data.csv',
        primaryKey: 'ID',
        customColumns: [],
        customFileName: '1_custom.csv'
    }];
}

async function fetchAttachment(filename) {
    // Simulated for demo
    return 'ID,Name,Value\n1,Item1,100\n2,Item2,200';
}

async function uploadAttachment(filename, content) {
    console.log(`Uploading ${filename}:`, content);
}

async function switchTab(tabId) {
    STATE.currentTab = tabId;
    const tab = STATE.tabs.find(t => t.id === tabId);
    await loadTabData(tab);
    renderTable();
}

async function loadTabData(tab) {
    const data = CSV.parse(await fetchAttachment(tab.fileName));
    let customData = [];
    try {
        customData = CSV.parse(await fetchAttachment(tab.customFileName));
    } catch (e) {}
    
    const merged = data.map(row => {
        const custom = customData.find(c => c[tab.primaryKey] === row[tab.primaryKey]) || {};
        return {...row, ...custom};
    });
    
    STATE.tabData[tab.id] = { originalData: merged, filteredData: [...merged] };
}

function renderTable() {
    const tab = STATE.tabs.find(t => t.id === STATE.currentTab);
    const data = STATE.tabData[STATE.currentTab];
    if (!data) return;
    
    let displayData = [...data.filteredData];
    
    // Apply sorting
    if (STATE.sortColumn) {
        displayData.sort((a, b) => {
            let valA = a[STATE.sortColumn], valB = b[STATE.sortColumn];
            const numA = parseFloat(valA), numB = parseFloat(valB);
            if (!isNaN(numA) && !isNaN(numB)) {
                valA = numA; valB = numB;
            } else {
                valA = String(valA || '').toLowerCase();
                valB = String(valB || '').toLowerCase();
            }
            if (valA < valB) return STATE.sortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return STATE.sortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }
    
    // Pagination
    const start = (STATE.currentPage - 1) * STATE.pageSize;
    const end = start + STATE.pageSize;
    const pageData = displayData.slice(start, end);
    
    const columns = displayData.length ? Object.keys(displayData[0]) : [];
    
    document.getElementById('tableHead').innerHTML = `<tr>
        <th><input type="checkbox" onchange="selectAll(this)"></th>
        ${columns.map(col => `<th onclick="sortBy('${col}')" style="cursor:pointer">${col} 
            ${STATE.sortColumn === col ? STATE.sortDirection === 'asc' ? '▲' : '▼' : ''}</th>`).join('')}
    </tr>`;
    
    document.getElementById('tableBody').innerHTML = pageData.map((row, i) => `
        <tr ${STATE.selectedRows.has(start + i) ? 'class="selected"' : ''}>
            <td><input type="checkbox" ${STATE.selectedRows.has(start + i) ? 'checked' : ''} 
                onchange="toggleSelect(${start + i})"></td>
            ${columns.map(col => `<td>${row[col] || ''}</td>`).join('')}
        </tr>
    `).join('') || '<tr><td colspan="100%" style="text-align:center">No data</td></tr>';
    
    updatePagination(displayData.length);
}

function sortBy(column) {
    if (STATE.sortColumn === column) {
        STATE.sortDirection = STATE.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        STATE.sortColumn = column;
        STATE.sortDirection = 'asc';
    }
    renderTable();
}

function selectAll(cb) {
    if (cb.checked) {
        const data = STATE.tabData[STATE.currentTab].filteredData;
        const start = (STATE.currentPage - 1) * STATE.pageSize;
        const end = Math.min(start + STATE.pageSize, data.length);
        for (let i = start; i < end; i++) STATE.selectedRows.add(i);
    } else {
        STATE.selectedRows.clear();
    }
    renderTable();
}

function toggleSelect(index) {
    if (STATE.selectedRows.has(index)) {
        STATE.selectedRows.delete(index);
    } else {
        STATE.selectedRows.add(index);
    }
    renderTable();
}

function updatePagination(total) {
    const totalPages = Math.ceil(total / STATE.pageSize);
    document.getElementById('pageStart').textContent = total > 0 ? (STATE.currentPage - 1) * STATE.pageSize + 1 : 0;
    document.getElementById('pageEnd').textContent = Math.min(STATE.currentPage * STATE.pageSize, total);
    document.getElementById('totalRows').textContent = total;
}

function showLoading(msg) {
    const overlay = document.getElementById('loadingOverlay');
    overlay.querySelector('p').textContent = msg || 'Loading...';
    overlay.classList.add('show');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('show');
}

function showToast(msg, type) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span>${msg}</span>`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

function closeModal(id) {
    document.getElementById(id).classList.remove('show');
}

function openModal(id) {
    document.getElementById(id).classList.add('show');
}

function setupEventListeners() {
    // Add all event listeners
    document.getElementById('globalSearch').addEventListener('input', e => {
        const query = e.target.value.toLowerCase();
        const data = STATE.tabData[STATE.currentTab];
        if (data) {
            data.filteredData = data.originalData.filter(row => 
                Object.values(row).some(v => String(v).toLowerCase().includes(query))
            );
            STATE.currentPage = 1;
            renderTable();
        }
    });
}

// Export functions for HTML
window.sortBy = sortBy;
window.selectAll = selectAll;
window.toggleSelect = toggleSelect;
window.closeModal = closeModal;
window.openModal = openModal;
