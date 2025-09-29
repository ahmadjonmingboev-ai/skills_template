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
    tableContainer.innerHTML = '<div class="error-state"><div class="empty-state-icon">⚠️</div><p style="font-size: 16px; margin-bottom: 8px;">Dashboard URL not configured</p><p style="font-size: 14px;">Please update the configuration in the JavaScript file.</p></div>';
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
  
  // Tableau options
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
  
  // Create viz
  try {
    vizObjects[tabId] = new tableau.Viz(hiddenDiv, config.url, options);
  } catch (error) {
    console.error('Error loading Tableau:', error);
    tableContainer.innerHTML = '<div class="error-state"><div class="empty-state-icon">❌</div><p style="font-size: 16px;">Error loading dashboard</p><p style="font-size: 14px; margin-top: 8px;">' + error.message + '</p></div>';
  }
}

// Extract data from Tableau
function extractData(tabId, viz) {
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
      tableContainer.innerHTML = '<div class="error-state"><div class="empty-state-icon">⚠️</div><p>No worksheet found in dashboard</p></div>';
      return;
    }
    
    // Get underlying data
    worksheet.getUnderlyingDataAsync().then(function(dataTable) {
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
      
      // Render table
      renderTable(tabId);
      
      // Clean up hidden viz
      const hiddenDiv = document.getElementById('hiddenViz' + containerNum);
      if (hiddenDiv) {
        vizObjects[tabId].dispose();
        hiddenDiv.remove();
      }
      
    }).catch(function(error) {
      console.error('Error extracting data:', error);
      tableContainer.innerHTML = '<div class="error-state"><div class="empty-state-icon">❌</div><p>Error extracting data from dashboard</p><p style="font-size: 14px; margin-top: 8px;">' + error.message + '</p></div>';
    });
    
  } catch (error) {
    console.error('Error:', error);
    tableContainer.innerHTML = '<div class="error-state"><div class="empty-state-icon">❌</div><p>Error processing dashboard</p></div>';
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
  
  // Headers with sorting
  data.columns.forEach(col => {
    const isSorted = sortColumn[tabId] === col;
    const sortClass = isSorted ? (sortDirection[tabId] === 'asc' ? 'sorted-asc' : 'sorted-desc') : '';
    tableHTML += `<th class="sortable ${sortClass}" onclick="sortTable('${tabId}', '${col}')">${col}`;
    tableHTML += `<input type="text" class="column-filter" placeholder="Filter..." onclick="event.stopPropagation()" oninput="filterColumn('${tabId}', '${col}', this.value)" value="${colFilters[col] || ''}">`;
    tableHTML += `</th>`;
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

// Filter by column
function filterColumn(tabId, column, value) {
  if (!columnFilters[tabId]) {
    columnFilters[tabId] = {};
  }
  columnFilters[tabId][column] = value;
  currentPage[tabId] = 1;
  renderTable(tabId);
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