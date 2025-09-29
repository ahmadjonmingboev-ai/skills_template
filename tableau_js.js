// Tableau Dashboard Configuration
// Replace these URLs with your actual Tableau Public dashboard URLs
const dashboardConfig = {
  tab1: 'YOUR_TABLEAU_PUBLIC_URL_1',
  tab2: 'YOUR_TABLEAU_PUBLIC_URL_2',
  tab3: 'YOUR_TABLEAU_PUBLIC_URL_3',
  tab4: 'YOUR_TABLEAU_PUBLIC_URL_4',
  tab5: 'YOUR_TABLEAU_PUBLIC_URL_5'
};

// Store Tableau viz objects
let vizObjects = {};
let loadedTabs = new Set();

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
  // Load Tableau JS API
  loadTableauAPI();
  
  // Setup tab click handlers
  setupTabHandlers();
  
  // Load first dashboard
  setTimeout(() => {
    loadDashboard('tab1');
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
      
      // Load dashboard if not already loaded
      if (!loadedTabs.has(tabId)) {
        loadDashboard(tabId);
      }
    });
  });
}

// Adjust container height based on Tableau dashboard size
function adjustContainerHeight(tabId, viz) {
  try {
    const workbook = viz.getWorkbook();
    const activeSheet = workbook.getActiveSheet();
    const sheetSize = activeSheet.getSize();
    
    // Get the container
    const containerDiv = document.getElementById('vizContainer' + tabId.replace('tab', ''));
    
    // Set minimum height based on sheet size if available
    if (sheetSize && sheetSize.maxSize) {
      const height = sheetSize.maxSize.height;
      if (height && height > 0) {
        containerDiv.style.minHeight = height + 'px';
      }
    }
  } catch (error) {
    console.log('Could not adjust height automatically:', error);
  }
}

// Load Tableau dashboard
function loadDashboard(tabId) {
  const containerDiv = document.getElementById('vizContainer' + tabId.replace('tab', ''));
  const url = dashboardConfig[tabId];
  
  if (!url || url.includes('YOUR_TABLEAU')) {
    containerDiv.innerHTML = '<div style="padding: 40px; text-align: center; color: #6b7280;"><p style="font-size: 16px; margin-bottom: 8px;">⚠️ Dashboard URL not configured</p><p style="font-size: 14px;">Please update the dashboard URL in the JavaScript configuration.</p></div>';
    return;
  }
  
  // Dispose existing viz if present
  if (vizObjects[tabId]) {
    vizObjects[tabId].dispose();
  }
  
  // Clear container
  containerDiv.innerHTML = '';
  
  // Tableau options
  const options = {
    hideTabs: true,
    hideToolbar: false,
    width: '100%',
    // Don't set height - let Tableau use its published dimensions
    onFirstInteractive: function() {
      console.log('Dashboard ' + tabId + ' loaded successfully');
      adjustContainerHeight(tabId, this);
    }
  };
  
  // Create new viz
  try {
    vizObjects[tabId] = new tableau.Viz(containerDiv, url, options);
    loadedTabs.add(tabId);
  } catch (error) {
    console.error('Error loading dashboard:', error);
    containerDiv.innerHTML = '<div style="padding: 40px; text-align: center; color: #ef4444;"><p style="font-size: 16px;">Error loading dashboard</p><p style="font-size: 14px; margin-top: 8px;">' + error.message + '</p></div>';
  }
}

// Cleanup function (optional - call if needed)
function disposeAllViz() {
  Object.keys(vizObjects).forEach(tabId => {
    if (vizObjects[tabId]) {
      vizObjects[tabId].dispose();
    }
  });
  vizObjects = {};
  loadedTabs.clear();
}