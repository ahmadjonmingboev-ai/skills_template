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
  
  // Load first tab
  setTimeout(() => {
    loadTableauData('tab1');
  }, 500);
  
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
      
      // Load data if not already loaded
      if (!loadedTabs.has(tabId)) {
        loadTableauData(tabId);
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

// Load Tableau data
function loadTableauData(tabId) {
  const config = dashboardConfig[tabId];
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
  if (!config || !config.url || config.url.includes('YOUR_TABLEAU')) {
    tableContainer.innerHTML = '<div class="loading-spinner"><p style="font-size: 18px;">⚙️</p><p>Dashboard URL not configured</p><p class="loading-hint">Please update the configuration in the JavaScript file</p></div>';
    return;
  }
  
  // Create hidden container for Tableau viz
  const hiddenDiv = document.createElement('div');
  hiddenDiv.style.position = 'absolute';
  hiddenDiv.style.left = '-9999px';
  hiddenDiv.style.width = '1px';
  hiddenDiv.style.height = '1px';
  hiddenDiv.id = 'hiddenViz' + containerNum;
  document.body.appendChild(hiddenDiv);
  
  // Tableau options with error handling
  const options = {
    hideTabs: true,
    hideToolbar: true,
    width: '800px',
    height: '600px',
    onFirstInteractive: function() {
      // Use the stored viz object instead of 'this'
      setTimeout(() => {
        extractData(tabId, vizObjects[tabId]);
      }, 100);
    }
  };
  
  // Monitor console errors for 401
  const originalFetch = window.fetch;
  const authCheckWrapper = async (...args) => {
    try {
      const response = await originalFetch(...args);
      if (response.status === 401 && args[0] && args[0].includes('tableau')) {
        console.error('Tableau 401 Unauthorized detected');
        showAuthRequired(tabId);
      }
      return response;
    } catch (error) {
      return originalFetch(...args);
    }
  };
  window.fetch = authCheckWrapper;
  
  // Create viz with error handling
  try {
    const viz = new tableau.Viz(hiddenDiv, config.url, options);
    vizObjects[tabId] = viz;
    
    // Listen for viz errors
    viz.addEventListener(tableau.TableauEventName.VIZ_RESIZE, function() {
      // Viz loaded successfully, restore fetch
      window.fetch = originalFetch;
    });
    
  } catch (error) {
    console.error('Error loading Tableau:', error);
    window.fetch = originalFetch;
    
    // Check if it's an auth error
    if (error.message && (error.message.includes('401') || 
                          error.message.toLowerCase().includes('unauthorized') ||
                          error.message.toLowerCase().includes('auth') || 
                          error.message.toLowerCase().includes('login') || 
                          error.message.toLowerCase().includes('credentials'))) {
      showAuthRequired(tabId);
    }
  }
}

// Extract data from Tableau
function extractData(tabId, viz, vizDiv, loadingOverlay) {
  const config = dashboardConfig[tabId];
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
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
      return;
    }
    
    // Get summary data (works with viewer permissions)
    worksheet.getSummaryDataAsync().then(function(dataTable) {
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
      
      // Clean up viz and overlay
      if (loadingOverlay && loadingOverlay.parentNode) {
        loadingOverlay.parentNode.removeChild(loadingOverlay);
      }
      
      if (vizDiv && vizDiv.parentNode) {
        viz.dispose();
        vizDiv.parentNode.removeChild(vizDiv);
      }
      
      // Clear table container and render table
      tableContainer.innerHTML = '';
      renderTable(tabId);
      
    }).catch(function(error) {
      console.error('Error extracting data:', error);
      tableContainer.innerHTML = '<div class="loading-spinner"><p style="font-size: 18px;">❌</p><p>Error extracting data</p><p class="loading-hint">' + error.message + '</p></div>';
    });
    
  } catch (error) {
    console.error('Error:', error);
    tableContainer.innerHTML = '<div class="loading-spinner"><p style="font-size: 18px;">❌</p><p>Error processing dashboard</p><p class="loading-hint">' + error.message + '</p></div>';
  }
}

// Retry loading data after authentication
function retryLoadData(tabId) {
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
  // Show loading message
  tableContainer.innerHTML = `
    <div class="loading-spinner">
      <div class="spinner-icon">⏳</div>
      <p>Loading data from Tableau...</p>
      <p class="loading-hint">If this takes too long, you may not have access to this dashboard</p>
    </div>
  `;
  
  // Retry loading
  loadTableauData(tabId);
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
      
      // Handle null/undefined
      if (valA == null) valA = '';
      if (valB == null) valB = '';
      
      // Try numeric comparison
      const numA = parseFloat(valA);
      const numB = parseFloat(valB);
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return dir === 'asc' ? numA - numB : numB - numA;
      }
      
      // String comparison
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
  
  // Headers with sorting (no filter inputs)
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
  
  // Previous button
  paginationHTML += `<button class="pagination-button" ${currentPage === 1 ? 'disabled' : ''} onclick="goToPage('${tabId}', ${currentPage - 1})">Previous</button>`;
  
  // Page numbers
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
  
  // Next button
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
    // Toggle direction
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
  
  // Copy current filters to temp
  tempFilters = JSON.parse(JSON.stringify(columnFilters[tabId] || {}));
  
  // Build filter form
  let formHTML = '';
  data.columns.forEach(col => {
    // Get unique values for this column
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
  
  // Clear all inputs in modal
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
  
  // Get filtered data
  let filteredRows = data.rows;
  const searchTerm = searchTerms[tabId];
  const colFilters = columnFilters[tabId];
  
  // Apply filters (same logic as render)
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
  
  // Create CSV content
  let csv = '';
  
  // Headers
  csv += data.columns.map(col => `"${col}"`).join(',') + '\n';
  
  // Rows
  filteredRows.forEach(row => {
    csv += data.columns.map(col => {
      const val = row[col] !== null && row[col] !== undefined ? row[col] : '';
      return `"${String(val).replace(/"/g, '""')}"`;
    }).join(',') + '\n';
  });
  
  // Download
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