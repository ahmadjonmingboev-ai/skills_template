// Tableau Dashboard Configuration
// Replace these with your actual Tableau Public dashboard URLs and worksheet names
const dashboardConfig = {
  tab1: {
    url: 'YOUR_TABLEAU_PUBLIC_URL_1',
    worksheetName: 'Sheet1' // The name of the worksheet/sheet in your dashboard
  },
  tab2: {
    url: 'YOUR_TABLEAU_PUBLIC_URL_2',
    worksheetName: 'Sheet1'
  },
  tab3: {
    url: 'YOUR_TABLEAU_PUBLIC_URL_3',
    worksheetName: 'Sheet1'
  },
  tab4: {
    url: 'YOUR_TABLEAU_PUBLIC_URL_4',
    worksheetName: 'Sheet1'
  },
  tab5: {
    url: 'YOUR_TABLEAU_PUBLIC_URL_5',
    worksheetName: 'Sheet1'
  }
};

// Store data and state
let tabData = {};
let currentPage = {};
let rowsPerPage = {};
let searchTerms = {};
let columnFilters = {};
let sortColumn = {};
let sortDirection = {};
let loadedTabs = new Set();
let vizObjects = {};
let currentFilterTab = null;
let tempFilters = {};

// Initialize
document.addEventListener('DOMContentLoaded', function() {
  loadTableauAPI();
  setupTabHandlers();
  setupControlHandlers();
  
  // Initialize state for all tabs
  for (let i = 1; i <= 5; i++) {
    const tabId = 'tab' + i;
    currentPage[tabId] = 1;
    rowsPerPage[tabId] = 25;
    searchTerms[tabId] = '';
    columnFilters[tabId] = {};
    sortColumn[tabId] = null;
    sortDirection[tabId] = 'asc';
  }
  
  // Show auth prompt for first tab
  showAuthPrompt('tab1');
  
  // Close modal on escape key
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      closeFilterModal();
    }
  });
  
  // Close modal on background click
  document.getElementById('filterModal').addEventListener('click', function(e) {
    if (e.target === this) {
      closeFilterModal();
    }
  });
});

// Load Tableau JS API
function loadTableauAPI() {
  if (!document.querySelector('script[src*="tableau"]')) {
    const script = document.createElement('script');
    script.src = 'https://public.tableau.com/javascripts/api/tableau-2.min.js';
    script.async = true;
    document.head.appendChild(script);
  }
}

// Setup tab click handlers
function setupTabHandlers() {
  const tabButtons = document.querySelectorAll('.tab-button');
  
  tabButtons.forEach(button => {
    button.addEventListener('click', function() {
      const tabId = this.getAttribute('data-tab');
      
      // Update active states
      tabButtons.forEach(btn => btn.classList.remove('active'));
      this.classList.add('active');
      
      // Show selected tab content
      const tabContents = document.querySelectorAll('.tab-content');
      tabContents.forEach(content => content.classList.remove('active'));
      document.getElementById(tabId).classList.add('active');
      
      // Show auth prompt if not loaded
      if (!loadedTabs.has(tabId)) {
        showAuthPrompt(tabId);
      }
    });
  });
}

// Setup control handlers (search, pagination)
function setupControlHandlers() {
  // Search inputs
  document.querySelectorAll('.search-input').forEach(input => {
    input.addEventListener('input', function() {
      const tabId = this.getAttribute('data-tab');
      searchTerms[tabId] = this.value.toLowerCase();
      currentPage[tabId] = 1;
      renderTable(tabId);
    });
  });
  
  // Rows per page selects
  document.querySelectorAll('.rows-per-page').forEach(select => {
    select.addEventListener('change', function() {
      const tabId = this.getAttribute('data-tab');
      rowsPerPage[tabId] = this.value === 'all' ? 'all' : parseInt(this.value);
      currentPage[tabId] = 1;
      renderTable(tabId);
    });
  });
}

// Show authentication prompt
function showAuthPrompt(tabId) {
  const config = dashboardConfig[tabId];
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
  if (!config || !config.url || config.url.includes('YOUR_TABLEAU')) {
    tableContainer.innerHTML = `
      <div class="auth-prompt">
        <div class="auth-icon">⚙️</div>
        <h3>Dashboard Not Configured</h3>
        <p>Please update the dashboard URL in the configuration</p>
      </div>
    `;
    return;
  }
  
  const authHTML = `
    <div class="auth-prompt">
      <div class="auth-icon">🔐</div>
      <h3>Authentication Required</h3>
      <p>To view this dashboard data, you need to authenticate with Tableau first.</p>
      <div class="auth-steps">
        <div class="auth-step">
          <span class="step-number">1</span>
          <span class="step-text">Click "Sign in to Tableau" below</span>
        </div>
        <div class="auth-step">
          <span class="step-number">2</span>
          <span class="step-text">Sign in with your credentials in the popup window</span>
        </div>
        <div class="auth-step">
          <span class="step-number">3</span>
          <span class="step-text">Close the popup and click "Load Data"</span>
        </div>
      </div>
      <div class="auth-buttons">
        <button class="auth-button primary" onclick="openTableauAuth('${tabId}')">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"></path>
            <polyline points="10 17 15 12 10 7"></polyline>
            <line x1="15" y1="12" x2="3" y2="12"></line>
          </svg>
          Sign in to Tableau
        </button>
        <button class="auth-button secondary" onclick="loadTableauData('${tabId}')">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="23 4 23 10 17 10"></polyline>
            <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path>
          </svg>
          Load Data
        </button>
      </div>
      <p class="auth-hint">Already signed in? Click "Load Data" directly.</p>
    </div>
  `;
  
  tableContainer.innerHTML = authHTML;
}

// Open Tableau for authentication
function openTableauAuth(tabId) {
  const config = dashboardConfig[tabId];
  if (config && config.url) {
    // Open in popup window
    const width = 1000;
    const height = 700;
    const left = (screen.width - width) / 2;
    const top = (screen.height - height) / 2;
    
    window.open(
      config.url, 
      'TableauAuth',
      `width=${width},height=${height},left=${left},top=${top},toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes`
    );
  }
}

// Load Tableau data - HIDDEN VIZ AFTER AUTH
function loadTableauData(tabId) {
  const config = dashboardConfig[tabId];
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
  // Show loading with water animation
  tableContainer.innerHTML = `
    <div class="loading-spinner">
      <div class="water-loader">
        <div class="water-fill"></div>
        <div class="water-wave"></div>
      </div>
      <p>Loading data from Tableau...</p>
      <p class="loading-hint">This may take a moment...</p>
    </div>
  `;
  
  // Create hidden container for Tableau viz
  const hiddenDiv = document.createElement('div');
  hiddenDiv.style.position = 'absolute';
  hiddenDiv.style.left = '-9999px';
  hiddenDiv.style.width = '1px';
  hiddenDiv.style.height = '1px';
  hiddenDiv.id = 'hiddenViz' + containerNum;
  document.body.appendChild(hiddenDiv);
  
  // Tableau options
  const options = {
    hideTabs: true,
    hideToolbar: true,
    width: '800px',
    height: '600px',
    onFirstInteractive: function() {
      console.log('Tableau viz loaded, extracting data...');
      
      // Update loading message
      tableContainer.innerHTML = `
        <div class="loading-spinner">
          <div class="water-loader">
            <div class="water-fill"></div>
            <div class="water-wave"></div>
          </div>
          <p>Extracting data...</p>
          <p class="loading-hint">Converting to table view...</p>
        </div>
      `;
      
      // Extract data
      setTimeout(() => {
        extractData(tabId, vizObjects[tabId], hiddenDiv, tableContainer);
      }, 500);
    }
  };
  
  // Create viz
  try {
    const viz = new tableau.Viz(hiddenDiv, config.url, options);
    vizObjects[tabId] = viz;
    
    // If it doesn't load in 10 seconds, show auth prompt again
    setTimeout(() => {
      if (!loadedTabs.has(tabId)) {
        // Clean up
        if (hiddenDiv && hiddenDiv.parentNode) {
          viz.dispose();
          hiddenDiv.remove();
        }
        
        tableContainer.innerHTML = `
          <div class="auth-prompt">
            <div class="auth-icon">⚠️</div>
            <h3>Authentication Required</h3>
            <p>Couldn't load data. You may need to sign in to Tableau.</p>
            <div class="auth-buttons">
              <button class="auth-button primary" onclick="openTableauAuth('${tabId}')">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"></path>
                  <polyline points="10 17 15 12 10 7"></polyline>
                  <line x1="15" y1="12" x2="3" y2="12"></line>
                </svg>
                Sign in to Tableau
              </button>
              <button class="auth-button secondary" onclick="loadTableauData('${tabId}')">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <polyline points="23 4 23 10 17 10"></polyline>
                  <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path>
                </svg>
                Try Again
              </button>
            </div>
          </div>
        `;
      }
    }, 10000);
    
  } catch (error) {
    console.error('Error loading Tableau:', error);
    
    // Clean up
    if (hiddenDiv && hiddenDiv.parentNode) {
      hiddenDiv.remove();
    }
    
    tableContainer.innerHTML = `
      <div class="auth-prompt">
        <div class="auth-icon">❌</div>
        <h3>Error Loading Dashboard</h3>
        <p>${error.message}</p>
        <div class="auth-buttons">
          <button class="auth-button primary" onclick="openTableauAuth('${tabId}')">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"></path>
              <polyline points="10 17 15 12 10 7"></polyline>
              <line x1="15" y1="12" x2="3" y2="12"></line>
            </svg>
            Sign in to Tableau
          </button>
          <button class="auth-button secondary" onclick="loadTableauData('${tabId}')">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="23 4 23 10 17 10"></polyline>
              <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path>
            </svg>
            Try Again
          </button>
        </div>
      </div>
    `;
  }
}

// Extract data from Tableau
function extractData(tabId, viz, hiddenDiv, tableContainer) {
  const config = dashboardConfig[tabId];
  
  try {
    const workbook = viz.getWorkbook();
    const worksheets = workbook.getActiveSheet().getWorksheets();
    
    // Find the worksheet by name or use first one
    let worksheet;
    if (config.worksheetName) {
      worksheet = worksheets.find(ws => ws.getName() === config.worksheetName);
    }
    if (!worksheet) {
      worksheet = worksheets[0];
    }
    
    if (!worksheet) {
      console.error('No worksheet found');
      tableContainer.innerHTML = '<div class="loading-spinner"><p style="font-size: 18px;">⚠️</p><p>No worksheet found in dashboard</p></div>';
      
      // Clean up
      if (hiddenDiv && hiddenDiv.parentNode) {
        viz.dispose();
        hiddenDiv.remove();
      }
      return;
    }
    
    // Get summary data
    const options = {
      maxRows: 0,
      ignoreAliases: false,
      ignoreSelection: true
    };
    
    worksheet.getSummaryDataAsync(options).then(
      function(dataTable) {
        const columns = dataTable.getColumns();
        const data = dataTable.getData();
        
        // Convert to array of objects
        const rows = data.map(row => {
          const obj = {};
          columns.forEach((col, idx) => {
            obj[col.getFieldName()] = row[idx].value;
          });
          return obj;
        });
        
        // Store data
        tabData[tabId] = {
          columns: columns.map(col => col.getFieldName()),
          rows: rows
        };
        
        loadedTabs.add(tabId);
        
        // Clean up
        if (hiddenDiv && hiddenDiv.parentNode) {
          viz.dispose();
          hiddenDiv.remove();
        }
        
        // Render table
        tableContainer.innerHTML = '';
        renderTable(tabId);
      },
      function(error) {
        console.error('Error extracting data:', error);
        tableContainer.innerHTML = '<div class="loading-spinner"><p style="font-size: 18px;">❌</p><p>Error extracting data</p><p class="loading-hint">' + error + '</p></div>';
        
        // Clean up
        if (hiddenDiv && hiddenDiv.parentNode) {
          viz.dispose();
          hiddenDiv.remove();
        }
      }
    );
    
  } catch (error) {
    console.error('Error:', error);
    tableContainer.innerHTML = '<div class="loading-spinner"><p style="font-size: 18px;">❌</p><p>Error processing dashboard</p><p class="loading-hint">' + error.message + '</p></div>';
    
    // Clean up
    if (hiddenDiv && hiddenDiv.parentNode) {
      viz.dispose();
      hiddenDiv.remove();
    }
  }
}

// Render table with pagination
function renderTable(tabId) {
  const data = tabData[tabId];
  if (!data) return;
  
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  const paginationContainer = document.getElementById('pagination' + containerNum);
  const recordCount = document.querySelector('#' + tabId + ' .record-count');
  
  // Filter data based on search
  let filteredRows = data.rows;
  const searchTerm = searchTerms[tabId];
  const colFilters = columnFilters[tabId];
  
  // Apply global search
  if (searchTerm) {
    filteredRows = filteredRows.filter(row => {
      return Object.values(row).some(val => 
        String(val).toLowerCase().includes(searchTerm)
      );
    });
  }
  
  // Apply column filters
  if (colFilters && Object.keys(colFilters).length > 0) {
    filteredRows = filteredRows.filter(row => {
      return Object.keys(colFilters).every(col => {
        const filterValue = colFilters[col].toLowerCase();
        if (!filterValue) return true;
        return String(row[col]).toLowerCase().includes(filterValue);
      });
    });
  }
  
  // Apply sorting
  if (sortColumn[tabId]) {
    const col = sortColumn[tabId];
    const dir = sortDirection[tabId];
    filteredRows = [...filteredRows].sort((a, b) => {
      let valA = a[col];
      let valB = b[col];
      
      if (valA == null) valA = '';
      if (valB == null) valB = '';
      
      const numA = parseFloat(valA);
      const numB = parseFloat(valB);
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return dir === 'asc' ? numA - numB : numB - numA;
      }
      
      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();
      
      if (dir === 'asc') {
        return strA < strB ? -1 : strA > strB ? 1 : 0;
      } else {
        return strA > strB ? -1 : strA < strB ? 1 : 0;
      }
    });
  }
  
  // Calculate pagination
  const totalRows = filteredRows.length;
  const isAllRows = rowsPerPage[tabId] === 'all';
  const perPage = isAllRows ? totalRows : rowsPerPage[tabId];
  const totalPages = isAllRows ? 1 : Math.ceil(totalRows / perPage);
  const page = Math.min(currentPage[tabId], totalPages || 1);
  currentPage[tabId] = page;
  
  const startIdx = isAllRows ? 0 : (page - 1) * perPage;
  const endIdx = isAllRows ? totalRows : Math.min(startIdx + perPage, totalRows);
  const pageRows = filteredRows.slice(startIdx, endIdx);
  
  // Update record count
  if (searchTerm || (colFilters && Object.keys(colFilters).length > 0)) {
    recordCount.textContent = `Showing ${startIdx + 1}-${endIdx} of ${totalRows} (filtered from ${data.rows.length} total)`;
  } else {
    recordCount.textContent = `Showing ${startIdx + 1}-${endIdx} of ${totalRows} records`;
  }
  
  // Build table
  if (pageRows.length === 0) {
    tableContainer.innerHTML = '<div class="empty-state"><div class="empty-state-icon">🔍</div><p>No records found</p></div>';
    paginationContainer.innerHTML = '';
    return;
  }
  
  let tableHTML = '<table class="data-table"><thead><tr>';
  
  // Headers with sorting
  data.columns.forEach(col => {
    const isSorted = sortColumn[tabId] === col;
    const sortClass = isSorted ? (sortDirection[tabId] === 'asc' ? 'sorted-asc' : 'sorted-desc') : '';
    tableHTML += `<th class="sortable ${sortClass}" onclick="sortTable('${tabId}', '${col}')">${col}</th>`;
  });
  tableHTML += '</tr></thead><tbody>';
  
  // Rows
  pageRows.forEach(row => {
    tableHTML += '<tr>';
    data.columns.forEach(col => {
      const value = row[col] !== null && row[col] !== undefined ? row[col] : '';
      tableHTML += `<td>${value}</td>`;
    });
    tableHTML += '</tr>';
  });
  
  tableHTML += '</tbody></table>';
  tableContainer.innerHTML = tableHTML;
  
  // Update filter badge
  updateFilterBadge(tabId);
  
  // Build pagination
  if (!isAllRows) {
    renderPagination(tabId, page, totalPages);
  } else {
    paginationContainer.innerHTML = '';
  }
}

// Render pagination controls
function renderPagination(tabId, currentPage, totalPages) {
  const containerNum = tabId.replace('tab', '');
  const paginationContainer = document.getElementById('pagination' + containerNum);
  
  if (totalPages <= 1) {
    paginationContainer.innerHTML = '';
    return;
  }
  
  let paginationHTML = '';
  
  paginationHTML += `<button class="pagination-button" ${currentPage === 1 ? 'disabled' : ''} onclick="goToPage('${tabId}', ${currentPage - 1})">Previous</button>`;
  
  const maxButtons = 5;
  let startPage = Math.max(1, currentPage - Math.floor(maxButtons / 2));
  let endPage = Math.min(totalPages, startPage + maxButtons - 1);
  
  if (endPage - startPage < maxButtons - 1) {
    startPage = Math.max(1, endPage - maxButtons + 1);
  }
  
  if (startPage > 1) {
    paginationHTML += `<button class="pagination-button" onclick="goToPage('${tabId}', 1)">1</button>`;
    if (startPage > 2) {
      paginationHTML += `<span class="pagination-info">...</span>`;
    }
  }
  
  for (let i = startPage; i <= endPage; i++) {
    paginationHTML += `<button class="pagination-button ${i === currentPage ? 'active' : ''}" onclick="goToPage('${tabId}', ${i})">${i}</button>`;
  }
  
  if (endPage < totalPages) {
    if (endPage < totalPages - 1) {
      paginationHTML += `<span class="pagination-info">...</span>`;
    }
    paginationHTML += `<button class="pagination-button" onclick="goToPage('${tabId}', ${totalPages})">${totalPages}</button>`;
  }
  
  paginationHTML += `<button class="pagination-button" ${currentPage === totalPages ? 'disabled' : ''} onclick="goToPage('${tabId}', ${currentPage + 1})">Next</button>`;
  
  paginationContainer.innerHTML = paginationHTML;
}

// Go to specific page
function goToPage(tabId, page) {
  currentPage[tabId] = page;
  renderTable(tabId);
}

// Sort table by column
function sortTable(tabId, column) {
  if (sortColumn[tabId] === column) {
    sortDirection[tabId] = sortDirection[tabId] === 'asc' ? 'desc' : 'asc';
  } else {
    sortColumn[tabId] = column;
    sortDirection[tabId] = 'asc';
  }
  renderTable(tabId);
}

// Update filter badge
function updateFilterBadge(tabId) {
  const containerNum = tabId.replace('tab', '');
  const badge = document.getElementById('filterBadge' + containerNum);
  const colFilters = columnFilters[tabId] || {};
  const activeFilters = Object.values(colFilters).filter(v => v && v.trim()).length;
  
  if (activeFilters > 0) {
    badge.textContent = activeFilters;
    badge.style.display = 'inline-block';
  } else {
    badge.style.display = 'none';
  }
}

// Open filter modal
function openFilterModal(tabId) {
  const data = tabData[tabId];
  if (!data) return;
  
  currentFilterTab = tabId;
  const modal = document.getElementById('filterModal');
  const modalBody = document.getElementById('filterModalBody');
  
  tempFilters = JSON.parse(JSON.stringify(columnFilters[tabId] || {}));
  
  let formHTML = '';
  data.columns.forEach(col => {
    const uniqueValues = [...new Set(data.rows.map(row => row[col]))].filter(v => v !== null && v !== undefined).sort();
    
    formHTML += `<div class="filter-group">`;
    formHTML += `<label>${col}</label>`;
    formHTML += `<div class="filter-select-wrapper">`;
    formHTML += `<select class="filter-select" id="filter_${col}" onchange="updateTempFilter('${col}', this.value)">`;
    formHTML += `<option value="">All</option>`;
    
    uniqueValues.slice(0, 100).forEach(val => {
      const selected = tempFilters[col] === String(val) ? 'selected' : '';
      formHTML += `<option value="${val}" ${selected}>${val}</option>`;
    });
    
    if (uniqueValues.length > 100) {
      formHTML += `<option disabled>... ${uniqueValues.length - 100} more values</option>`;
    }
    
    formHTML += `</select>`;
    formHTML += `</div>`;
    formHTML += `<input type="text" class="filter-input" placeholder="Or type to filter..." value="${tempFilters[col] || ''}" oninput="updateTempFilter('${col}', this.value)">`;
    formHTML += `</div>`;
  });
  
  modalBody.innerHTML = formHTML;
  modal.classList.add('active');
}

// Close filter modal
function closeFilterModal() {
  const modal = document.getElementById('filterModal');
  modal.classList.remove('active');
  currentFilterTab = null;
  tempFilters = {};
}

// Update temp filter
function updateTempFilter(column, value) {
  tempFilters[column] = value;
}

// Apply filters
function applyFilters() {
  if (!currentFilterTab) return;
  
  columnFilters[currentFilterTab] = JSON.parse(JSON.stringify(tempFilters));
  currentPage[currentFilterTab] = 1;
  renderTable(currentFilterTab);
  closeFilterModal();
}

// Clear all filters
function clearAllFilters() {
  if (!currentFilterTab) return;
  
  tempFilters = {};
  columnFilters[currentFilterTab] = {};
  
  document.querySelectorAll('.filter-select').forEach(select => select.value = '');
  document.querySelectorAll('.filter-input').forEach(input => input.value = '');
}

// Export data to CSV
function exportData(tabId) {
  const data = tabData[tabId];
  if (!data) {
    alert('No data to export');
    return;
  }
  
  let filteredRows = data.rows;
  const searchTerm = searchTerms[tabId];
  const colFilters = columnFilters[tabId];
  
  if (searchTerm) {
    filteredRows = filteredRows.filter(row => {
      return Object.values(row).some(val => 
        String(val).toLowerCase().includes(searchTerm)
      );
    });
  }
  
  if (colFilters && Object.keys(colFilters).length > 0) {
    filteredRows = filteredRows.filter(row => {
      return Object.keys(colFilters).every(col => {
        const filterValue = colFilters[col].toLowerCase();
        if (!filterValue) return true;
        return String(row[col]).toLowerCase().includes(filterValue);
      });
    });
  }
  
  if (sortColumn[tabId]) {
    const col = sortColumn[tabId];
    const dir = sortDirection[tabId];
    filteredRows = [...filteredRows].sort((a, b) => {
      let valA = a[col];
      let valB = b[col];
      
      if (valA == null) valA = '';
      if (valB == null) valB = '';
      
      const numA = parseFloat(valA);
      const numB = parseFloat(valB);
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return dir === 'asc' ? numA - numB : numB - numA;
      }
      
      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();
      
      if (dir === 'asc') {
        return strA < strB ? -1 : strA > strB ? 1 : 0;
      } else {
        return strA > strB ? -1 : strA < strB ? 1 : 0;
      }
    });
  }
  
  let csv = '';
  csv += data.columns.map(col => `"${col}"`).join(',') + '\n';
  
  filteredRows.forEach(row => {
    csv += data.columns.map(col => {
      const val = row[col] !== null && row[col] !== undefined ? row[col] : '';
      return `"${String(val).replace(/"/g, '""')}"`;
    }).join(',') + '\n';
  });
  
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  
  const containerNum = tabId.replace('tab', '');
  const filename = `tableau_report_${containerNum}_${new Date().toISOString().slice(0, 10)}.csv`;
  
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}