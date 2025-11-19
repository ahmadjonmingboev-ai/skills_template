// Tableau Dashboard Configuration
// Replace these with your actual Tableau Server dashboard URLs and worksheet names
const dashboardConfig = {
  tab1: {
    url: 'YOUR_TABLEAU_SERVER_URL_1',
    worksheetName: 'Sheet1', // The name of the worksheet/sheet in your dashboard
    useMethod: 'summary', // Use 'summary' for getSummaryDataAsync (no row limit, but gets viz data) or 'underlying' for getUnderlyingDataAsync (true underlying data, 10K limit)
    formatColumns: {
      // Specify columns to format with decimal precision
      // Example: 'Amount': { decimals: 2 }, 'Quantity': { decimals: 0 }
    }
  },
  tab2: {
    url: 'YOUR_TABLEAU_SERVER_URL_2',
    worksheetName: 'Sheet1',
    useMethod: 'underlying', // Recommended for reports with < 10K rows
    formatColumns: {}
  },
  tab3: {
    url: 'YOUR_TABLEAU_SERVER_URL_3',
    worksheetName: 'Sheet1',
    useMethod: 'underlying',
    formatColumns: {}
  },
  tab4: {
    url: 'YOUR_TABLEAU_SERVER_URL_4',
    worksheetName: 'Sheet1',
    useMethod: 'underlying',
    formatColumns: {}
  },
  tab5: {
    url: 'YOUR_TABLEAU_SERVER_URL_5',
    worksheetName: 'Sheet1',
    useMethod: 'underlying',
    formatColumns: {}
  }
};

// Sign-in Dashboard URL - This dashboard will be shown in the sign-in modal
// Users will authenticate through this embedded viz
const SIGNIN_DASHBOARD_URL = 'YOUR_TABLEAU_DASHBOARD_URL_FOR_SIGNIN';

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
let columnOrder = {}; // Store custom column order per tab
let tableauAPIReady = false; // Track if Tableau API is loaded
let loadingTimers = {}; // Track loading progress timers per tab
let errorTimers = {}; // Track error display timers per tab

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
    columnOrder[tabId] = null; // Will be populated when data loads
  }
  
  // Show loading progress for tab1 immediately (before API loads)
  showLoadingProgress('tab1');
  
  // Show sign-in wizard modal after short delay
  setTimeout(() => {
    showSignInWizard();
  }, 1000);
  
  // Also start trying to load data immediately (will wait for API if needed)
  setTimeout(() => {
    console.log('Attempting to start data load from DOMContentLoaded...');
    checkAndLoadFirstTab();
  }, 500);
  
  // Close modal on escape key
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      closeFilterModal();
      closeSignInModal();
    }
  });
  
  // Close modals on background click
  document.getElementById('filterModal').addEventListener('click', function(e) {
    if (e.target === this) {
      closeFilterModal();
    }
  });
  
  document.getElementById('signInModal').addEventListener('click', function(e) {
    if (e.target === this) {
      closeSignInModal();
    }
  });
});

// Load Tableau JS API
function loadTableauAPI() {
  if (!document.querySelector('script[src*="tableau"]')) {
    const script = document.createElement('script');
    script.src = 'https://public.tableau.com/javascripts/api/tableau-2.min.js';
    script.async = true;
    script.onload = function() {
      console.log('Tableau API loaded successfully');
      tableauAPIReady = true;
      // Check if we need to load first tab after API is ready
      checkAndLoadFirstTab();
    };
    script.onerror = function() {
      console.error('Failed to load Tableau API');
      tableauAPIReady = false;
    };
    document.head.appendChild(script);
  } else {
    // Script already exists, check if tableau is available
    if (typeof tableau !== 'undefined') {
      tableauAPIReady = true;
      console.log('Tableau API already available');
      checkAndLoadFirstTab(); // Load first tab immediately
    } else {
      // Wait for tableau to be available
      console.log('Waiting for Tableau API to be available...');
      const checkInterval = setInterval(() => {
        if (typeof tableau !== 'undefined') {
          tableauAPIReady = true;
          console.log('Tableau API now available');
          clearInterval(checkInterval);
          checkAndLoadFirstTab();
        }
      }, 100);
    }
  }
}

// Check and load first tab after API is ready
let loadAttempts = 0;
const maxLoadAttempts = 50; // 10 seconds max (50 * 200ms)

function checkAndLoadFirstTab() {
  console.log('checkAndLoadFirstTab called (attempt ' + (loadAttempts + 1) + '), tableauAPIReady:', tableauAPIReady);
  
  if (tableauAPIReady && typeof tableau !== 'undefined') {
    console.log('✓ Tableau API ready - Loading tab1 data...');
    loadTableauData('tab1');
    loadAttempts = 0; // Reset counter
  } else {
    loadAttempts++;
    if (loadAttempts < maxLoadAttempts) {
      console.log('⏳ Tableau API not ready yet, waiting... (attempt ' + loadAttempts + '/' + maxLoadAttempts + ')');
      setTimeout(checkAndLoadFirstTab, 200);
    } else {
      console.error('❌ Tableau API failed to load after ' + maxLoadAttempts + ' attempts');
      loadAttempts = 0;
    }
  }
}

// Show sign-in wizard modal on page load
function showSignInWizard() {
  // Check if user recently attempted sign-in (within 4 hours)
  const signInAttempted = localStorage.getItem('tableau_sign_in_attempted');
  if (signInAttempted) {
    const attemptTime = parseInt(signInAttempted);
    const fourHoursInMs = 4 * 60 * 60 * 1000; // 4 hours in milliseconds
    const now = Date.now();
    
    if (now - attemptTime < fourHoursInMs) {
      console.log('Sign-in attempted recently (within 4 hours), skipping wizard');
      return;
    } else {
      console.log('Sign-in attempt expired (>4 hours), clearing flag');
      localStorage.removeItem('tableau_sign_in_attempted');
    }
  }
  
  // Don't show if already signed in and data is loading
  if (loadedTabs.size > 0) {
    console.log('Data already loaded, skipping sign-in wizard');
    return;
  }
  
  const modal = document.getElementById('signInModal');
  const vizContainer = document.getElementById('signInVizContainer');
  const statusEl = document.getElementById('signInStatus');
  
  // Show wizard-style message
  modal.classList.add('active');
  statusEl.textContent = 'Are you signed in to Tableau Server?';
  statusEl.style.color = '#374151';
  statusEl.style.fontSize = '16px';
  statusEl.style.fontWeight = '500';
  
  // Create Yes/No buttons
  vizContainer.innerHTML = `
    <div style="display: flex; flex-direction: column; align-items: center; gap: 20px; padding: 40px;">
      <p style="font-size: 14px; color: #6b7280; margin: 0;">If you're not signed in, the dashboards won't load properly.</p>
      <div style="display: flex; gap: 16px;">
        <button class="btn-primary" onclick="handleSignInWizardYes()" style="padding: 12px 32px; font-size: 15px;">
          Yes, I'm Signed In
        </button>
        <button class="btn-secondary" onclick="handleSignInWizardNo()" style="padding: 12px 32px; font-size: 15px;">
          No, Sign Me In
        </button>
      </div>
    </div>
  `;
}

// Handle "Yes, I'm Signed In" - just close and let it load
function handleSignInWizardYes() {
  console.log('User confirmed they are signed in');
  closeSignInModal();
  // Force load data if not started
  if (!loadedTabs.has('tab1')) {
    console.log('Starting data load after wizard confirmation');
    setTimeout(() => {
      if (tableauAPIReady) {
        loadTableauData('tab1');
      } else {
        // Wait for API
        const waitForAPI = setInterval(() => {
          if (tableauAPIReady) {
            clearInterval(waitForAPI);
            loadTableauData('tab1');
          }
        }, 200);
      }
    }, 100);
  }
}

// Handle "No, Sign Me In" - trigger sign-in flow
function handleSignInWizardNo() {
  console.log('User requested sign-in');
  
  // Set localStorage flag with current timestamp (valid for 4 hours)
  localStorage.setItem('tableau_sign_in_attempted', Date.now().toString());
  console.log('Set sign-in attempt flag in localStorage');
  
  closeSignInModal();
  // Small delay then open actual sign-in
  setTimeout(() => {
    openSignInModal();
  }, 300);
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

// Format numeric value based on configuration
function formatValue(value, columnName, tabId) {
  const config = dashboardConfig[tabId];
  
  // Check if this column has specific formatting config
  if (config.formatColumns && config.formatColumns[columnName]) {
    const format = config.formatColumns[columnName];
    const num = parseFloat(value);
    if (!isNaN(num)) {
      return num.toFixed(format.decimals);
    }
  }
  
  // Auto-detect currency/amount columns and format to 2 decimals
  const lowerColName = columnName.toLowerCase();
  if (lowerColName.includes('amount') || 
      lowerColName.includes('price') || 
      lowerColName.includes('cost') || 
      lowerColName.includes('revenue') ||
      lowerColName.includes('total') ||
      lowerColName.includes('balance')) {
    const num = parseFloat(value);
    if (!isNaN(num) && value.toString().includes('.')) {
      return num.toFixed(2);
    }
  }
  
  return value;
}

// Show loading with animated progress bar
function showLoadingProgress(tabId) {
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
  tableContainer.innerHTML = `
    <div class="loading-spinner">
      <div class="spinner-icon">⏳</div>
      <p>Loading data from Tableau...</p>
      <div class="progress-bar-container">
        <div class="progress-bar" id="progressBar${containerNum}"></div>
      </div>
      <p class="loading-hint" id="progressText${containerNum}">0%</p>
    </div>
  `;
  
  // Animate progress bar (simulated progress)
  let progress = 0;
  const progressBar = document.getElementById('progressBar' + containerNum);
  const progressText = document.getElementById('progressText' + containerNum);
  
  // Clear any existing timer
  if (loadingTimers[tabId]) {
    clearInterval(loadingTimers[tabId]);
  }
  
  // Animate progress (fast at first, then slower)
  loadingTimers[tabId] = setInterval(() => {
    if (progress < 30) {
      progress += 2;
    } else if (progress < 60) {
      progress += 1;
    } else if (progress < 90) {
      progress += 0.5;
    } else if (progress < 95) {
      progress += 0.2;
    }
    
    if (progressBar && progressText) {
      progressBar.style.width = progress + '%';
      progressText.textContent = Math.round(progress) + '%';
    }
  }, 100);
}

// Complete loading progress
function completeLoadingProgress(tabId) {
  const containerNum = tabId.replace('tab', '');
  const progressBar = document.getElementById('progressBar' + containerNum);
  const progressText = document.getElementById('progressText' + containerNum);
  
  // Clear timer
  if (loadingTimers[tabId]) {
    clearInterval(loadingTimers[tabId]);
    delete loadingTimers[tabId];
  }
  
  // Complete to 100%
  if (progressBar && progressText) {
    progressBar.style.width = '100%';
    progressText.textContent = '100%';
  }
}

// Load Tableau data
function loadTableauData(tabId) {
  const config = dashboardConfig[tabId];
  const containerNum = tabId.replace('tab', '');
  const tableContainer = document.getElementById('tableContainer' + containerNum);
  
  if (!config || !config.url || config.url.includes('YOUR_TABLEAU')) {
    tableContainer.innerHTML = '<div class="loading-spinner"><div class="spinner-icon">⚙️</div><p>Dashboard URL not configured</p><p class="loading-hint">Please update the configuration in the JavaScript file</p></div>';
    return;
  }
  
  // Check if Tableau API is ready
  if (!tableauAPIReady || typeof tableau === 'undefined') {
    console.log('Waiting for Tableau API to load...');
    tableContainer.innerHTML = '<div class="loading-spinner"><div class="spinner-icon">⏳</div><p>Loading Tableau API...</p></div>';
    // Retry after a short delay
    setTimeout(() => {
      loadTableauData(tabId);
    }, 500);
    return;
  }
  
  // Show loading with progress bar
  showLoadingProgress(tabId);
  
  // Create hidden container for Tableau viz
  // Make it tiny but technically visible to allow auth handshake
  const hiddenDiv = document.createElement('div');
  hiddenDiv.style.position = 'absolute';
  hiddenDiv.style.top = '-9999px';
  hiddenDiv.style.left = '0';
  hiddenDiv.style.width = '1px';
  hiddenDiv.style.height = '1px';
  hiddenDiv.style.opacity = '0';
  hiddenDiv.style.pointerEvents = 'none';
  hiddenDiv.style.overflow = 'hidden';
  hiddenDiv.id = 'hiddenViz' + containerNum;
  document.body.appendChild(hiddenDiv);
  
  // Tableau options
  const options = {
    hideTabs: true,
    hideToolbar: true,
    width: '800px',
    height: '600px',
    onFirstInteractive: function() {
      console.log('Tableau viz loaded successfully for ' + tabId);
      // Use the stored viz object instead of 'this'
      setTimeout(() => {
        extractData(tabId, vizObjects[tabId]);
      }, 100);
    },
    onFirstVizSizeKnown: function() {
      console.log('Tableau viz size known for ' + tabId);
    }
  };
  
  // Create viz
  try {
    console.log('Initializing Tableau viz for ' + tabId + ' with URL:', config.url);
    vizObjects[tabId] = new tableau.Viz(hiddenDiv, config.url, options);
  } catch (error) {
    // Log to console for debugging but don't show to user (viz often loads successfully anyway)
    console.error('Error initializing Tableau viz for ' + tabId + ':', error);
    // Don't display error message - let loading animation continue and data will appear when ready
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
      console.error('No worksheet found for ' + tabId);
      tableContainer.innerHTML = '<div class="loading-spinner"><div class="spinner-icon">❌</div><p>Worksheet not found</p><p class="loading-hint">Check the worksheetName in configuration</p></div>';
      return;
    }
    
    console.log('Extracting data from worksheet:', worksheet.getName());
    
    // Check which method to use based on config
    const useMethod = config.useMethod || 'underlying'; // Default to underlying if not specified
    
    if (useMethod === 'summary') {
      console.log('Using getSummaryDataAsync for ' + tabId + ' (no row limit)');
      
      // Get summary data (works with viewer permissions, no row limit)
      worksheet.getSummaryDataAsync().then(function(dataTable) {
        console.log('Summary data extracted successfully for ' + tabId + ': ' + dataTable.getData().length + ' rows');
        const columns = dataTable.getColumns();
        const data = dataTable.getData();
        
        processAndStoreData(tabId, columns, data, containerNum);
        
      }).catch(function(error) {
        handleDataExtractionError(tabId, error, tableContainer);
      });
      
    } else {
      console.log('Using getUnderlyingDataAsync for ' + tabId + ' (true underlying data, 10K server limit)');
      
      // Get underlying data (requires "Download Full Data" permission for viewers)
      // Note: Server may have a hard limit (typically 10,000 rows)
      worksheet.getUnderlyingDataAsync().then(function(dataTable) {
        const rowCount = dataTable.getData().length;
        console.log('Underlying data extracted successfully for ' + tabId + ': ' + rowCount + ' rows');
        
        // Warn if we hit the limit
        if (rowCount === 10000) {
          console.warn('Warning: Received exactly 10,000 rows. This may be the server limit. Actual data might have more rows.');
        }
        
        const columns = dataTable.getColumns();
        const data = dataTable.getData();
        
        processAndStoreData(tabId, columns, data, containerNum);
        
      }).catch(function(error) {
        handleDataExtractionError(tabId, error, tableContainer);
      });
    }
    
  } catch (error) {
    // Log to console for debugging but don't show to user (data often loads successfully anyway)
    console.error('Error in extractData for ' + tabId + ':', error);
    // Don't display error message - let loading animation continue and data will appear when ready
  }
}

// Process and store extracted data
function processAndStoreData(tabId, columns, data, containerNum) {
  // Clear any error timers since data loaded successfully
  if (errorTimers[tabId]) {
    clearTimeout(errorTimers[tabId]);
    delete errorTimers[tabId];
  }
  
  // Complete progress bar
  completeLoadingProgress(tabId);
  
  // Clear sign-in attempt flag since data loaded successfully
  // This means user is authenticated and data is working
  if (localStorage.getItem('tableau_sign_in_attempted')) {
    console.log('Data loaded successfully, clearing sign-in attempt flag');
    localStorage.removeItem('tableau_sign_in_attempted');
  }
  
  // Get column names
  const columnNames = columns.map(col => col.getFieldName());
  
  // Convert to array of objects with RAW values (no formatting yet)
  // Formatting will be done at render time for better performance
  const rows = data.map(row => {
    const obj = {};
    columns.forEach((col, idx) => {
      const colName = col.getFieldName();
      obj[colName] = row[idx].value;
    });
    return obj;
  });
  
  // Store data
  tabData[tabId] = {
    columns: columnNames,
    rows: rows
  };
  
  // Initialize column order if not set (use original order from Tableau)
  if (!columnOrder[tabId]) {
    columnOrder[tabId] = [...columnNames];
  }
  
  loadedTabs.add(tabId);
  
  console.log('Data processing complete for ' + tabId + ': ' + rows.length + ' total rows');
  
  // Render table
  renderTable(tabId);
  
  // Clean up hidden viz
  const hiddenDiv = document.getElementById('hiddenViz' + containerNum);
  if (hiddenDiv) {
    vizObjects[tabId].dispose();
    hiddenDiv.remove();
  }
}

// Handle data extraction errors (with delay to avoid showing transient errors)
function handleDataExtractionError(tabId, error, tableContainer) {
  console.error('Error extracting data for ' + tabId + ':', error);
  
  // Don't show error immediately - wait 8 seconds to see if data loads anyway
  // Many "then is not a function" type errors resolve themselves
  errorTimers[tabId] = setTimeout(() => {
    // Only show error if we still don't have data after 8 seconds
    if (!tabData[tabId] || !tabData[tabId].rows || tabData[tabId].rows.length === 0) {
      // Stop progress animation
      completeLoadingProgress(tabId);
      
      // Show helpful error message based on error type
      let errorMessage = 'Failed to extract data from Tableau';
      let errorHint = error.message || 'Unknown error';
      
      if (error.message && error.message.includes('permission')) {
        errorHint = 'You may not have "Download Full Data" permission. Please contact your Tableau admin.';
      } else if (error.message && (error.message.includes('401') || error.message.includes('unauthorized'))) {
        errorHint = 'Authentication required. Please use the "Sign In" button and refresh the page.';
      }
      
      tableContainer.innerHTML = '<div class="loading-spinner"><div class="spinner-icon">❌</div><p>' + errorMessage + '</p><p class="loading-hint">' + errorHint + '</p></div>';
    }
  }, 8000); // Wait 8 seconds before showing error
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
        const filterValue = colFilters[col];
        if (!filterValue) return true;
        
        // Handle comma-separated values (multiple selections)
        const filterValues = filterValue.split(',').map(v => v.trim().toLowerCase());
        const rowValue = String(row[col]).toLowerCase();
        
        // Check if row value matches any of the filter values
        return filterValues.some(fv => rowValue.includes(fv));
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
  
  // Use custom column order if set, otherwise use original order
  const displayColumns = columnOrder[tabId] || data.columns;
  
  let tableHTML = '<table class="data-table" id="dataTable' + containerNum + '"><thead><tr>';
  
  // Headers with sorting and drag-drop (resize removed)
  displayColumns.forEach((col, colIndex) => {
    const isSorted = sortColumn[tabId] === col;
    const sortClass = isSorted ? (sortDirection[tabId] === 'asc' ? 'sorted-asc' : 'sorted-desc') : '';
    const colId = 'col_' + tabId + '_' + colIndex;
    
    tableHTML += `<th class="sortable ${sortClass}" 
                      id="${colId}"
                      draggable="true"
                      ondragstart="handleColumnDragStart(event, '${tabId}', ${colIndex})"
                      ondragover="handleColumnDragOver(event)"
                      ondrop="handleColumnDrop(event, '${tabId}', ${colIndex})"
                      ondragend="handleColumnDragEnd(event)"
                      onclick="sortTable('${tabId}', '${col}')">
                    ${col}
                  </th>`;
  });
  tableHTML += '</tr></thead><tbody>';
  
  // Rows (using custom column order) - Apply formatting here for performance
  pageRows.forEach(row => {
    tableHTML += '<tr>';
    displayColumns.forEach(col => {
      const rawValue = row[col];
      const formattedValue = rawValue !== null && rawValue !== undefined ? formatValue(rawValue, col, tabId) : '';
      tableHTML += `<td>${formattedValue}</td>`;
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

// Column Reordering Functionality
let draggedColumnIndex = null;
let draggedTabId = null;

function handleColumnDragStart(event, tabId, colIndex) {
  draggedColumnIndex = colIndex;
  draggedTabId = tabId;
  event.target.style.opacity = '0.5';
  event.dataTransfer.effectAllowed = 'move';
  event.dataTransfer.setData('text/html', event.target.innerHTML);
}

function handleColumnDragOver(event) {
  if (event.preventDefault) {
    event.preventDefault();
  }
  event.dataTransfer.dropEffect = 'move';
  
  // Add visual feedback
  const th = event.target.closest('th');
  if (th) {
    th.style.borderLeft = '3px solid #3b82f6';
  }
  
  return false;
}

function handleColumnDrop(event, tabId, targetColIndex) {
  if (event.stopPropagation) {
    event.stopPropagation();
  }
  
  if (draggedTabId !== tabId || draggedColumnIndex === null) {
    return false;
  }
  
  // Reorder columns
  if (draggedColumnIndex !== targetColIndex) {
    const columns = [...columnOrder[tabId]];
    const draggedColumn = columns[draggedColumnIndex];
    columns.splice(draggedColumnIndex, 1);
    columns.splice(targetColIndex, 0, draggedColumn);
    columnOrder[tabId] = columns;
    
    // Re-render table with new column order
    renderTable(tabId);
  }
  
  return false;
}

function handleColumnDragEnd(event) {
  event.target.style.opacity = '';
  
  // Remove all border highlights
  document.querySelectorAll('th').forEach(th => {
    th.style.borderLeft = '';
  });
  
  draggedColumnIndex = null;
  draggedTabId = null;
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
  
  // Build filter form with grid layout
  let formHTML = '<div class="filters-grid">';
  
  data.columns.forEach(col => {
    // Get unique values for this column
    const uniqueValues = [...new Set(data.rows.map(row => row[col]))].filter(v => v !== null && v !== undefined).sort();
    
    // Parse existing filter - could be single value or comma-separated
    const existingFilter = tempFilters[col] || '';
    const selectedValues = existingFilter ? existingFilter.split(',').map(v => v.trim()) : [];
    
    const hasFilter = existingFilter.trim();
    const filterClass = hasFilter ? 'has-filter' : '';
    
    formHTML += `<div class="filter-group ${filterClass}" id="filterGroup_${col}">`;
    formHTML += `<label>`;
    formHTML += `<span>${col}</span>`;
    if (hasFilter) {
      formHTML += `<span class="filter-active-badge">Active</span>`;
    }
    formHTML += `</label>`;
    
    formHTML += `<div class="filter-select-wrapper">`;
    formHTML += `<select class="filter-select" multiple id="filter_select_${col}" onchange="updateTempFilterFromSelect('${col}', this)">`;
    
    uniqueValues.slice(0, 100).forEach(val => {
      const selected = selectedValues.includes(String(val)) ? 'selected' : '';
      const displayVal = String(val).length > 50 ? String(val).substring(0, 50) + '...' : val;
      formHTML += `<option value="${val}" ${selected}>${displayVal}</option>`;
    });
    
    if (uniqueValues.length > 100) {
      formHTML += `<option disabled>... ${uniqueValues.length - 100} more values (use text filter)</option>`;
    }
    
    formHTML += `</select>`;
    formHTML += `</div>`;
    formHTML += `<div class="filter-input-wrapper">`;
    formHTML += `<input type="text" class="filter-input" placeholder="Or type values (comma-separated)..." value="${tempFilters[col] || ''}" oninput="updateTempFilterFromInput('${col}', this.value)">`;
    formHTML += `</div>`;
    formHTML += `</div>`;
  });
  
  formHTML += '</div>';
  modalBody.innerHTML = formHTML;
  
  // Update active filters count
  updateActiveFiltersCount();
  
  modal.classList.add('active');
}

// Update temp filter from select (multiple)
function updateTempFilterFromSelect(column, selectElement) {
  const selectedOptions = Array.from(selectElement.selectedOptions).map(opt => opt.value);
  const value = selectedOptions.join(', ');
  
  tempFilters[column] = value;
  updateFilterGroupHighlight(column);
  updateActiveFiltersCount();
  
  // Update the text input too
  const textInput = document.querySelector(`input.filter-input[oninput*="${column}"]`);
  if (textInput) {
    textInput.value = value;
  }
}

// Update temp filter from input
function updateTempFilterFromInput(column, value) {
  tempFilters[column] = value;
  updateFilterGroupHighlight(column);
  updateActiveFiltersCount();
  
  // Update the select - handle comma-separated values
  const select = document.getElementById('filter_select_' + column);
  if (select) {
    // Clear all selections first
    Array.from(select.options).forEach(opt => opt.selected = false);
    
    // Select matching options
    if (value) {
      const filterValues = value.split(',').map(v => v.trim());
      Array.from(select.options).forEach(opt => {
        if (filterValues.includes(opt.value)) {
          opt.selected = true;
        }
      });
    }
  }
}

// Update filter group highlight
function updateFilterGroupHighlight(column) {
  const group = document.getElementById('filterGroup_' + column);
  if (!group) return;
  
  const hasFilter = tempFilters[column] && tempFilters[column].trim();
  
  if (hasFilter) {
    group.classList.add('has-filter');
    // Update or add badge
    const label = group.querySelector('label');
    let badge = label.querySelector('.filter-active-badge');
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'filter-active-badge';
      badge.textContent = 'Active';
      label.appendChild(badge);
    }
  } else {
    group.classList.remove('has-filter');
    // Remove badge
    const badge = group.querySelector('.filter-active-badge');
    if (badge) {
      badge.remove();
    }
  }
}

// Update active filters count
function updateActiveFiltersCount() {
  const count = Object.values(tempFilters).filter(v => v && v.trim()).length;
  const countEl = document.getElementById('activeFiltersCount');
  if (countEl) {
    countEl.textContent = count === 0 ? 'No filters active' : count === 1 ? '1 filter active' : `${count} filters active`;
  }
}

// Close filter modal
function closeFilterModal() {
  const modal = document.getElementById('filterModal');
  modal.classList.remove('active');
  currentFilterTab = null;
  tempFilters = {};
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
  
  // Clear all inputs and selects in modal
  document.querySelectorAll('.filter-select').forEach(select => select.value = '');
  document.querySelectorAll('.filter-input').forEach(input => input.value = '');
  document.querySelectorAll('.filter-group').forEach(group => group.classList.remove('has-filter'));
  document.querySelectorAll('.filter-active-badge').forEach(badge => badge.remove());
  
  updateActiveFiltersCount();
}

// Sign In Modal Functions
let signInViz = null;

function openSignInModal() {
  const modal = document.getElementById('signInModal');
  const vizContainer = document.getElementById('signInVizContainer');
  const countdownEl = document.getElementById('countdown');
  const statusEl = document.getElementById('signInStatus');
  
  // Check if dashboard URL is configured
  if (!SIGNIN_DASHBOARD_URL || SIGNIN_DASHBOARD_URL.includes('YOUR_TABLEAU')) {
    statusEl.textContent = 'Sign-in dashboard URL not configured. Please update SIGNIN_DASHBOARD_URL in the configuration.';
    statusEl.style.color = '#ef4444';
    modal.classList.add('active');
    return;
  }
  
  // Show modal
  modal.classList.add('active');
  statusEl.textContent = 'Loading Tableau dashboard for authentication...';
  statusEl.style.color = '#6b7280';
  
  // Embed Tableau viz for authentication
  vizContainer.innerHTML = '<div id="signInVizEmbed"></div>';
  
  const options = {
    hideTabs: true,
    hideToolbar: false,
    width: '100%',
    height: '500px',
    onFirstInteractive: function() {
      console.log('Sign-in viz loaded successfully');
      statusEl.textContent = 'Please sign in to Tableau Server if prompted. Page will refresh automatically.';
      statusEl.style.color = '#10b981';
      
      // Start countdown after viz is loaded
      startSignInCountdown();
    }
  };
  
  try {
    const embedDiv = document.getElementById('signInVizEmbed');
    signInViz = new tableau.Viz(embedDiv, SIGNIN_DASHBOARD_URL, options);
  } catch (error) {
    console.error('Error loading sign-in viz:', error);
    statusEl.textContent = 'Failed to load authentication dashboard: ' + error.message;
    statusEl.style.color = '#ef4444';
  }
}

function startSignInCountdown() {
  const countdownEl = document.getElementById('countdown');
  const modal = document.getElementById('signInModal');
  
  // Start countdown (15 seconds)
  let seconds = 15;
  countdownEl.textContent = seconds;
  
  const countdown = setInterval(() => {
    seconds--;
    countdownEl.textContent = seconds;
    
    if (seconds <= 0) {
      clearInterval(countdown);
      
      // Dispose viz before reload
      if (signInViz) {
        signInViz.dispose();
        signInViz = null;
      }
      
      location.reload();
    }
  }, 1000);
  
  // Store countdown ID in case user closes modal
  modal.dataset.countdownId = countdown;
}

function closeSignInModal() {
  const modal = document.getElementById('signInModal');
  const countdownId = modal.dataset.countdownId;
  const vizContainer = document.getElementById('signInVizContainer');
  
  // Clear countdown if exists
  if (countdownId) {
    clearInterval(parseInt(countdownId));
  }
  
  // Dispose viz if exists
  if (signInViz) {
    signInViz.dispose();
    signInViz = null;
  }
  
  // Clear container
  vizContainer.innerHTML = '';
  
  modal.classList.remove('active');
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
        const filterValue = colFilters[col];
        if (!filterValue) return true;
        
        // Handle comma-separated values (multiple selections)
        const filterValues = filterValue.split(',').map(v => v.trim().toLowerCase());
        const rowValue = String(row[col]).toLowerCase();
        
        // Check if row value matches any of the filter values
        return filterValues.some(fv => rowValue.includes(fv));
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
  
  // Rows (with formatting applied)
  filteredRows.forEach(row => {
    csv += data.columns.map(col => {
      const rawVal = row[col];
      const formattedVal = rawVal !== null && rawVal !== undefined ? formatValue(rawVal, col, tabId) : '';
      return `"${String(formattedVal).replace(/"/g, '""')}"`;
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