// ==========================================
// Configuration Constants
// ==========================================

// Confluence Configuration (UPDATE THESE)
const CONFLUENCE_BASE_URL = 'mysite/wiki';
const CONFLUENCE_PAGE_ID = '1235678990';
const CSV_FILENAME_CMPS = 'cmp_library.csv';
const CSV_FILENAME_AUDIT = 'cmp_audit_log.csv';
const CSV_FILENAME_ANALYTICS = 'cmp_order_analytics.csv';

// Pagination
const CMPS_PER_PAGE = 12;

// Auto-refresh interval (milliseconds) - Set to 0 to disable
const AUTO_REFRESH_INTERVAL = 5 * 60 * 1000; // 5 minutes

// Search debounce delay
const SEARCH_DEBOUNCE_DELAY = 300;

// Toast duration
const TOAST_DURATION = 5000;
const CARD_STAGGER_DELAY = 50;

// Category colors
const CATEGORY_COLORS = {
  'VPN': '#305edb',
  'SSO': '#2a6b3c',
  'API': '#fab728',
  'Database': '#9a231a',
  'Tool': '#6366f1',
  'Security': '#0f1632',
  'Network': '#14b8a6',
  'Other': '#64748b'
};

// Region colors
const REGION_COLORS = {
  'US': '#305edb',
  'EU': '#2a6b3c',
  'APAC': '#fab728',
  'Global': '#64748b',
  'All': '#26292b'
};

// Default categories and regions
const DEFAULT_CATEGORIES = ['VPN', 'SSO', 'API', 'Database', 'Tool', 'Security', 'Network', 'Other'];
const DEFAULT_REGIONS = ['US', 'EU', 'APAC', 'Global', 'All'];

// ==========================================
// Global State Variables
// ==========================================

let allCmps = [];
let filteredCmps = [];
let currentPage = 1;
let currentView = 'grid';
let currentSort = 'alphabetical';
let currentCategory = '';
let currentRegion = '';
let searchQuery = '';
let editMode = false;
let deleteMode = false;
let currentUserName = '';
let dynamicCategories = [];
let dynamicRegions = [];
let searchDebounceTimer = null;
let cmpToDelete = null;
let autoRefreshTimer = null;
let analyticsRecords = [];

// ==========================================
// Utility Functions
// ==========================================

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function escapeCSVField(field) {
  if (field == null) return '';
  const str = String(field);
  
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  
  return str;
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const next = line[i + 1];
    
    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  result.push(current.trim());
  return result;
}

function parseCSV(csvText) {
  if (!csvText || !csvText.trim()) return [];
  
  const lines = csvText.split('\n').filter(line => line.trim());
  if (lines.length === 0) return [];
  
  const headers = parseCSVLine(lines[0]);
  const data = [];
  
  for (let i = 1; i < lines.length; i++) {
    const values = parseCSVLine(lines[i]);
    const obj = {};
    
    headers.forEach((header, index) => {
      obj[header] = values[index] || '';
    });
    
    data.push(obj);
  }
  
  return data;
}

function convertToCSV(data, headers) {
  if (!data || data.length === 0) return '';
  
  const csvRows = [];
  csvRows.push(headers.join(','));
  
  for (const row of data) {
    const values = headers.map(header => escapeCSVField(row[header]));
    csvRows.push(values.join(','));
  }
  
  return csvRows.join('\n');
}

function formatDate(dateString) {
  if (!dateString) return '-';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}

function generateId() {
  return 'cmp-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
}

function getCategoryColor(category) {
  return CATEGORY_COLORS[category] || CATEGORY_COLORS['Other'];
}

function getRegionColor(region) {
  return REGION_COLORS[region] || REGION_COLORS['All'];
}

function debounce(func, delay) {
  return function(...args) {
    clearTimeout(searchDebounceTimer);
    searchDebounceTimer = setTimeout(() => func.apply(this, args), delay);
  };
}

// ==========================================
// Toast Notifications
// ==========================================

function showToast(type, message) {
  const container = document.getElementById('toastContainer');
  
  const titleMap = {
    'success': 'Success!',
    'error': 'Error!',
    'warning': 'Warning!',
    'info': 'Information'
  };
  
  const iconMap = {
    'success': 'check-circle',
    'error': 'x-circle',
    'warning': 'alert-circle',
    'info': 'alert-circle'
  };
  
  const toast = document.createElement('div');
  toast.className = `cmp-toast cmp-toast-${type}`;
  toast.innerHTML = `
    <svg class="cmp-toast-icon cmp-icon"><use href="#icon-${iconMap[type]}"></use></svg>
    <div class="cmp-toast-content">
      <h4 class="cmp-toast-title">${titleMap[type]}</h4>
      <p class="cmp-toast-message">${escapeHtml(message)}</p>
    </div>
    <button class="cmp-toast-close" onclick="this.parentElement.remove()">
      <svg class="cmp-icon cmp-icon-sm"><use href="#icon-close"></use></svg>
    </button>
  `;
  
  container.appendChild(toast);
  
  setTimeout(() => {
    toast.classList.add('removing');
    setTimeout(() => toast.remove(), 300);
  }, TOAST_DURATION);
}

// ==========================================
// Confluence API Functions
// ==========================================

async function loadCurrentUser() {
  try {
    const response = await fetch(`${CONFLUENCE_BASE_URL}/rest/api/user/current`);
    const data = await response.json();
    currentUserName = data.displayName || 'Unknown User';
  } catch (error) {
    console.error('Error loading current user:', error);
    currentUserName = 'Unknown User';
  }
}

async function fetchCSVFromConfluence(filename) {
  try {
    console.log('=== FETCHING CSV ===');
    console.log('Filename:', filename);
    
    const attachmentsUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment`;
    console.log('Attachments URL:', attachmentsUrl);
    
    const response = await fetch(attachmentsUrl, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      },
      credentials: 'include'
    });
    
    console.log('Attachments response status:', response.status);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error('Failed to fetch attachments:', errorText);
      throw new Error(`Failed to fetch attachments: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    console.log('Attachments found:', data.results.length);
    console.log('Attachment titles:', data.results.map(a => a.title));
    
    const attachment = data.results.find(att => att.title === filename);
    
    if (!attachment) {
      console.log(`${filename} not found - first time setup`);
      return '';
    }
    
    console.log('Found attachment:', attachment.title);
    
    const downloadUrl = `${CONFLUENCE_BASE_URL}${attachment._links.download}`;
    console.log('Downloading from:', downloadUrl);
    
    const downloadResponse = await fetch(downloadUrl, {
      method: 'GET',
      credentials: 'include'
    });
    
    if (!downloadResponse.ok) {
      throw new Error(`Failed to download file: ${downloadResponse.status}`);
    }
    
    const csvText = await downloadResponse.text();
    console.log('Downloaded CSV size:', csvText.length, 'characters');
    
    return csvText;
  } catch (error) {
    console.error(`Error fetching ${filename}:`, error);
    return '';
  }
}

async function uploadCSVToConfluence(filename, content) {
  try {
    console.log('=== UPLOAD STARTING ===');
    console.log('Filename:', filename);
    console.log('Content size:', content.length, 'bytes');
    
    const formData = new FormData();
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8' });
    formData.append('file', blob, filename);
    formData.append('comment', `Updated via CMP Library - ${new Date().toISOString()}`);
    
    console.log('Blob created. Size:', blob.size, 'bytes');
    
    const attachmentsUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment`;
    console.log('Fetching attachments from:', attachmentsUrl);
    
    const attachmentsResponse = await fetch(attachmentsUrl, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'include'
    });
    
    console.log('Attachments response status:', attachmentsResponse.status);
    
    if (!attachmentsResponse.ok) {
      const errorText = await attachmentsResponse.text();
      console.error('Failed to fetch attachments:', errorText);
      throw new Error(`Failed to fetch attachments: ${attachmentsResponse.status} ${attachmentsResponse.statusText}`);
    }
    
    const attachmentsData = await attachmentsResponse.json();
    console.log('Total attachments found:', attachmentsData.results.length);
    
    const existingAttachment = attachmentsData.results.find(att => att.title === filename);
    console.log('Existing attachment found:', !!existingAttachment);
    
    let uploadUrl;
    if (existingAttachment) {
      uploadUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment/${existingAttachment.id}/data`;
      console.log('Updating existing attachment. ID:', existingAttachment.id);
    } else {
      uploadUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment`;
      console.log('Creating new attachment');
    }
    
    console.log('Upload URL:', uploadUrl);
    console.log('Starting upload...');
    
    const uploadResponse = await fetch(uploadUrl, {
      method: 'POST',
      headers: { 'X-Atlassian-Token': 'no-check' },
      body: formData,
      credentials: 'include'
    });
    
    console.log('Upload response status:', uploadResponse.status);
    console.log('Upload response statusText:', uploadResponse.statusText);
    
    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text();
      console.error('Upload failed with response:', errorText);
      throw new Error(`Failed to upload file: ${uploadResponse.status} ${uploadResponse.statusText} - ${errorText}`);
    }
    
    const result = await uploadResponse.json();
    console.log('=== UPLOAD SUCCESS ===');
    console.log('Upload result:', result);
    
    if (result.results && result.results[0]) {
      console.log('Uploaded file info:', {
        id: result.results[0].id,
        title: result.results[0].title,
        fileSize: result.results[0].extensions?.fileSize
      });
    }
    
    return result;
  } catch (error) {
    console.error('=== UPLOAD ERROR ===');
    console.error('Error:', error.message);
    console.error('Full error:', error);
    throw error;
  }
}

// ==========================================
// Data Loading & Processing
// ==========================================

async function loadCmps() {
  try {
    console.log('=== LOADING CMPS ===');
    showSkeleton();
    
    console.log('Fetching CSV from Confluence...');
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_CMPS);
    console.log('CSV data received. Length:', csvData.length);
    
    if (!csvData) {
      console.log('No CSV data found - first time setup');
      allCmps = [];
    } else {
      console.log('Parsing CSV...');
      allCmps = parseCSV(csvData);
      console.log('Parsed CMPs count:', allCmps.length);
      
      if (allCmps.length > 0) {
        console.log('Sample CMP:', {
          id: allCmps[0].CMP_ID,
          name: allCmps[0].CMP_Name,
          category: allCmps[0].CMP_Category
        });
      }
    }
    
    console.log('Processing CMPs...');
    processCmps();
    
    console.log('Applying filters and displaying...');
    applyFiltersAndDisplay();
    
    hideSkeleton();
    console.log('=== LOADING COMPLETE ===');
    
  } catch (error) {
    console.error('=== LOADING ERROR ===');
    console.error('Error loading CMPs:', error);
    showToast('error', 'Failed to load CMPs. Please check console for details.');
    hideSkeleton();
    document.getElementById('emptyState').style.display = 'block';
  }
}

function processCmps() {
  extractCategoriesAndRegions();
  buildChips();
}

function extractCategoriesAndRegions() {
  const categoriesSet = new Set(DEFAULT_CATEGORIES);
  const regionsSet = new Set(DEFAULT_REGIONS);
  
  allCmps.forEach(cmp => {
    if (cmp.CMP_Category) categoriesSet.add(cmp.CMP_Category);
    if (cmp.CMP_Region) regionsSet.add(cmp.CMP_Region);
  });
  
  dynamicCategories = Array.from(categoriesSet).sort();
  dynamicRegions = Array.from(regionsSet).sort();
  
  populateDataLists();
}

function populateDataLists() {
  const categoryList = document.getElementById('categoryList');
  const regionList = document.getElementById('regionList');
  
  categoryList.innerHTML = '';
  regionList.innerHTML = '';
  
  dynamicCategories.forEach(category => {
    const option = document.createElement('option');
    option.value = category;
    categoryList.appendChild(option);
  });
  
  dynamicRegions.forEach(region => {
    const option = document.createElement('option');
    option.value = region;
    regionList.appendChild(option);
  });
}

function buildChips() {
  const categoryChips = document.getElementById('categoryChips');
  const regionChips = document.getElementById('regionChips');
  
  // Build category chips
  const categoryCounts = {};
  allCmps.forEach(cmp => {
    const cat = cmp.CMP_Category || 'Other';
    categoryCounts[cat] = (categoryCounts[cat] || 0) + 1;
  });
  
  const totalCmps = allCmps.length;
  
  let categoryHtml = `
    <div class="cmp-chip ${currentCategory === '' ? 'active' : ''}" onclick="selectCategoryChip('')">
      All <span class="cmp-chip-count">(${totalCmps})</span>
    </div>
  `;
  
  dynamicCategories.forEach(category => {
    const count = categoryCounts[category] || 0;
    if (count > 0) {
      categoryHtml += `
        <div class="cmp-chip ${currentCategory === category ? 'active' : ''}" onclick="selectCategoryChip('${escapeHtml(category)}')">
          ${escapeHtml(category)} <span class="cmp-chip-count">(${count})</span>
        </div>
      `;
    }
  });
  
  categoryChips.innerHTML = categoryHtml;
  
  // Build region chips
  const regionCounts = {};
  allCmps.forEach(cmp => {
    const reg = cmp.CMP_Region || 'All';
    regionCounts[reg] = (regionCounts[reg] || 0) + 1;
  });
  
  let regionHtml = `
    <div class="cmp-chip ${currentRegion === '' ? 'active' : ''}" onclick="selectRegionChip('')">
      All <span class="cmp-chip-count">(${totalCmps})</span>
    </div>
  `;
  
  dynamicRegions.forEach(region => {
    const count = regionCounts[region] || 0;
    if (count > 0) {
      regionHtml += `
        <div class="cmp-chip ${currentRegion === region ? 'active' : ''}" onclick="selectRegionChip('${escapeHtml(region)}')">
          ${escapeHtml(region)} <span class="cmp-chip-count">(${count})</span>
        </div>
      `;
    }
  });
  
  regionChips.innerHTML = regionHtml;
}

// ==========================================
// Filtering & Sorting
// ==========================================

function selectCategoryChip(category) {
  currentCategory = category;
  buildChips();
  applyFiltersAndDisplay();
}

function selectRegionChip(region) {
  currentRegion = region;
  buildChips();
  applyFiltersAndDisplay();
}

function applyFiltersAndDisplay() {
  filteredCmps = allCmps.filter(cmp => {
    if (currentCategory && cmp.CMP_Category !== currentCategory) return false;
    if (currentRegion && cmp.CMP_Region !== currentRegion) return false;
    
    if (searchQuery) {
      const query = searchQuery.toLowerCase();
      const matchesName = cmp.CMP_Name?.toLowerCase().includes(query);
      const matchesDescription = cmp.CMP_Description?.toLowerCase().includes(query);
      const matchesOwner = cmp.CMP_Owner?.toLowerCase().includes(query);
      const matchesCategory = cmp.CMP_Category?.toLowerCase().includes(query);
      const matchesRegion = cmp.CMP_Region?.toLowerCase().includes(query);
      
      if (!matchesName && !matchesDescription && !matchesOwner && !matchesCategory && !matchesRegion) {
        return false;
      }
    }
    
    return true;
  });
  
  sortCmps();
  updateResultsCount();
  updateClearFiltersButton();
  
  currentPage = 1;
  displayCmps();
}

function sortCmps() {
  switch (currentSort) {
    case 'alphabetical':
      filteredCmps.sort((a, b) => (a.CMP_Name || '').localeCompare(b.CMP_Name || ''));
      break;
    case 'alphabetical-desc':
      filteredCmps.sort((a, b) => (b.CMP_Name || '').localeCompare(a.CMP_Name || ''));
      break;
    case 'newest':
      filteredCmps.sort((a, b) => new Date(b.CMP_Date_Added || 0) - new Date(a.CMP_Date_Added || 0));
      break;
    case 'oldest':
      filteredCmps.sort((a, b) => new Date(a.CMP_Date_Added || 0) - new Date(b.CMP_Date_Added || 0));
      break;
  }
}

function updateResultsCount() {
  const resultsEl = document.getElementById('resultsCount');
  const total = filteredCmps.length;
  const start = (currentPage - 1) * CMPS_PER_PAGE + 1;
  const end = Math.min(currentPage * CMPS_PER_PAGE, total);
  
  if (total === 0) {
    resultsEl.textContent = 'No CMPs found';
  } else {
    resultsEl.textContent = `Showing ${start}-${end} of ${total} CMP${total !== 1 ? 's' : ''}`;
  }
}

function updateClearFiltersButton() {
  const btn = document.getElementById('clearFiltersBtn');
  const hasFilters = searchQuery || currentCategory || currentRegion;
  btn.style.display = hasFilters ? 'block' : 'none';
}

function clearFilters() {
  searchQuery = '';
  currentCategory = '';
  currentRegion = '';
  
  document.getElementById('searchInput').value = '';
  
  buildChips();
  applyFiltersAndDisplay();
}

// ==========================================
// Display Functions
// ==========================================

function displayCmps() {
  if (currentView === 'grid') {
    displayGridView();
  } else {
    displayListView();
  }
  
  updatePagination();
}

function displayGridView() {
  const container = document.getElementById('cmpGrid');
  const emptyState = document.getElementById('emptyState');
  
  container.style.display = 'grid';
  document.getElementById('cmpList').style.display = 'none';
  
  if (filteredCmps.length === 0) {
    container.style.display = 'none';
    emptyState.style.display = 'block';
    return;
  }
  
  emptyState.style.display = 'none';
  
  const start = (currentPage - 1) * CMPS_PER_PAGE;
  const end = start + CMPS_PER_PAGE;
  const cmpsToShow = filteredCmps.slice(start, end);
  
  container.innerHTML = cmpsToShow.map((cmp, index) => createCmpCard(cmp, index)).join('');
}

function createCmpCard(cmp, index) {
  const categoryColor = getCategoryColor(cmp.CMP_Category);
  const regionColor = getRegionColor(cmp.CMP_Region);
  const hasTemplate = cmp.CMP_Template_URL && cmp.CMP_Template_URL.trim();
  const hasAttachment = cmp.CMP_Attachment_Code && cmp.CMP_Attachment_Code.trim();
  
  const iconHtml = cmp.CMP_Icon 
    ? `<img src="${escapeHtml(cmp.CMP_Icon)}" alt="${escapeHtml(cmp.CMP_Name)}" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" />
       <svg class="cmp-icon" style="display: none; color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`
    : `<svg class="cmp-icon" style="color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`;
  
  return `
    <div class="cmp-card" style="animation-delay: ${index * CARD_STAGGER_DELAY}ms">
      ${cmp.CMP_Tag ? `<div class="cmp-badge-tag">${escapeHtml(cmp.CMP_Tag)}</div>` : ''}
      
      <div class="cmp-card-header">
        <div class="cmp-app-icon" style="background: ${categoryColor}20;">
          ${iconHtml}
        </div>
        <div class="cmp-card-title-section">
          <h3 class="cmp-card-title">${escapeHtml(cmp.CMP_Name)}</h3>
          <div class="cmp-card-badges">
            <span class="cmp-badge" style="background: ${categoryColor};">
              ${escapeHtml(cmp.CMP_Category)}
            </span>
            <span class="cmp-badge" style="background: ${regionColor};">
              ${escapeHtml(cmp.CMP_Region)}
            </span>
          </div>
        </div>
      </div>
      
      <p class="cmp-card-description">${escapeHtml(cmp.CMP_Description || 'No description available')}</p>
      
      <div class="cmp-card-meta">
        ${cmp.CMP_Owner ? `
          <div class="cmp-meta-item">
            <svg class="cmp-icon cmp-icon-xs"><use href="#icon-user"></use></svg>
            <span>${escapeHtml(cmp.CMP_Owner)}</span>
          </div>
        ` : ''}
        ${cmp.CMP_Processing_Time ? `
          ${cmp.CMP_Owner ? '<div class="cmp-meta-separator"></div>' : ''}
          <div class="cmp-meta-item">
            <svg class="cmp-icon cmp-icon-xs"><use href="#icon-clock"></use></svg>
            <span>${escapeHtml(cmp.CMP_Processing_Time)}</span>
          </div>
        ` : ''}
        ${cmp.CMP_Prerequisites ? `
          ${(cmp.CMP_Owner || cmp.CMP_Processing_Time) ? '<div class="cmp-meta-separator"></div>' : ''}
          <div class="cmp-meta-item">
            <svg class="cmp-icon cmp-icon-xs"><use href="#icon-alert-circle"></use></svg>
            <span>${escapeHtml(cmp.CMP_Prerequisites)}</span>
          </div>
        ` : ''}
      </div>
      
      <div class="cmp-card-actions">
        ${hasTemplate 
          ? `<button class="cmp-btn cmp-btn-primary cmp-btn-sm" onclick="openOrderModal('${cmp.CMP_ID}')">
               <svg class="cmp-icon cmp-icon-xs"><use href="#icon-order"></use></svg>
               Order
             </button>`
          : `<button class="cmp-btn cmp-btn-na cmp-btn-sm" disabled>
               N/A
             </button>`
        }
        
        ${hasTemplate 
          ? `<button class="cmp-btn cmp-btn-secondary cmp-btn-sm" onclick="copyTemplateLink('${cmp.CMP_ID}')">
               <svg class="cmp-icon cmp-icon-xs"><use href="#icon-copy"></use></svg>
               Copy
             </button>`
          : `<button class="cmp-btn cmp-btn-na cmp-btn-sm" disabled>
               N/A
             </button>`
        }
        
        ${hasAttachment 
          ? `<button class="cmp-btn cmp-btn-secondary cmp-btn-sm" onclick="openAttachmentModal('${cmp.CMP_ID}')">
               <svg class="cmp-icon cmp-icon-xs"><use href="#icon-attachment"></use></svg>
               View
             </button>`
          : `<button class="cmp-btn cmp-btn-na cmp-btn-sm" disabled>
               N/A
             </button>`
        }
      </div>
      
      ${editMode || deleteMode ? `
        <div class="cmp-card-edit-actions">
          ${editMode ? `
            <button class="cmp-icon-btn" onclick="editCmp('${cmp.CMP_ID}')" title="Edit">
              <svg class="cmp-icon cmp-icon-sm"><use href="#icon-edit"></use></svg>
            </button>
          ` : ''}
          ${deleteMode ? `
            <button class="cmp-icon-btn danger" onclick="deleteCmp('${cmp.CMP_ID}', '${escapeHtml(cmp.CMP_Name)}')" title="Delete">
              <svg class="cmp-icon cmp-icon-sm"><use href="#icon-trash"></use></svg>
            </button>
          ` : ''}
        </div>
      ` : ''}
    </div>
  `;
}

function displayListView() {
  const container = document.getElementById('cmpList');
  const emptyState = document.getElementById('emptyState');
  
  container.style.display = 'block';
  document.getElementById('cmpGrid').style.display = 'none';
  
  if (filteredCmps.length === 0) {
    container.style.display = 'none';
    emptyState.style.display = 'block';
    return;
  }
  
  emptyState.style.display = 'none';
  
  const start = (currentPage - 1) * CMPS_PER_PAGE;
  const end = start + CMPS_PER_PAGE;
  const cmpsToShow = filteredCmps.slice(start, end);
  
  container.innerHTML = cmpsToShow.map(cmp => createListItem(cmp)).join('');
}

function createListItem(cmp) {
  const categoryColor = getCategoryColor(cmp.CMP_Category);
  const regionColor = getRegionColor(cmp.CMP_Region);
  const hasTemplate = cmp.CMP_Template_URL && cmp.CMP_Template_URL.trim();
  const hasAttachment = cmp.CMP_Attachment_Code && cmp.CMP_Attachment_Code.trim();
  
  const iconHtml = cmp.CMP_Icon 
    ? `<img src="${escapeHtml(cmp.CMP_Icon)}" alt="${escapeHtml(cmp.CMP_Name)}" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" />
       <svg class="cmp-icon" style="display: none; color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`
    : `<svg class="cmp-icon" style="color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`;
  
  return `
    <div class="cmp-list-item">
      <div class="cmp-list-icon" style="background: ${categoryColor}20;">
        ${iconHtml}
      </div>
      
      <div class="cmp-list-content">
        <h4 class="cmp-list-title">${escapeHtml(cmp.CMP_Name)}</h4>
        <p class="cmp-list-description">${escapeHtml(cmp.CMP_Description || 'No description')}</p>
      </div>
      
      <div class="cmp-list-badges">
        ${cmp.CMP_Tag ? `<div class="cmp-badge-tag">${escapeHtml(cmp.CMP_Tag)}</div>` : ''}
        <span class="cmp-badge" style="background: ${categoryColor};">
          ${escapeHtml(cmp.CMP_Category)}
        </span>
        <span class="cmp-badge" style="background: ${regionColor};">
          ${escapeHtml(cmp.CMP_Region)}
        </span>
      </div>
      
      <div class="cmp-list-actions">
        ${hasTemplate ? `
          <button class="cmp-btn cmp-btn-primary cmp-btn-sm" onclick="openOrderModal('${cmp.CMP_ID}')">
            <svg class="cmp-icon cmp-icon-xs"><use href="#icon-order"></use></svg>
          </button>
        ` : ''}
        ${hasTemplate ? `
          <button class="cmp-btn cmp-btn-secondary cmp-btn-sm" onclick="copyTemplateLink('${cmp.CMP_ID}')">
            <svg class="cmp-icon cmp-icon-xs"><use href="#icon-copy"></use></svg>
          </button>
        ` : ''}
        ${hasAttachment ? `
          <button class="cmp-btn cmp-btn-secondary cmp-btn-sm" onclick="openAttachmentModal('${cmp.CMP_ID}')">
            <svg class="cmp-icon cmp-icon-xs"><use href="#icon-attachment"></use></svg>
          </button>
        ` : ''}
        ${editMode ? `
          <button class="cmp-icon-btn" onclick="editCmp('${cmp.CMP_ID}')" title="Edit">
            <svg class="cmp-icon cmp-icon-sm"><use href="#icon-edit"></use></svg>
          </button>
        ` : ''}
        ${deleteMode ? `
          <button class="cmp-icon-btn danger" onclick="deleteCmp('${cmp.CMP_ID}', '${escapeHtml(cmp.CMP_Name)}')" title="Delete">
            <svg class="cmp-icon cmp-icon-sm"><use href="#icon-trash"></use></svg>
          </button>
        ` : ''}
      </div>
    </div>
  `;
}

// ==========================================
// Pagination
// ==========================================

function updatePagination() {
  const container = document.getElementById('pagination');
  const totalPages = Math.ceil(filteredCmps.length / CMPS_PER_PAGE);
  
  if (totalPages <= 1) {
    container.style.display = 'none';
    return;
  }
  
  container.style.display = 'flex';
  
  let paginationHTML = `
    <button class="cmp-pagination-btn" onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>
      Previous
    </button>
  `;
  
  for (let i = 1; i <= Math.min(totalPages, 5); i++) {
    paginationHTML += `
      <button class="cmp-page-btn ${i === currentPage ? 'active' : ''}" onclick="goToPage(${i})">
        ${i}
      </button>
    `;
  }
  
  if (totalPages > 5) {
    paginationHTML += '<span style="padding: 8px;">...</span>';
    paginationHTML += `
      <button class="cmp-page-btn ${totalPages === currentPage ? 'active' : ''}" onclick="goToPage(${totalPages})">
        ${totalPages}
      </button>
    `;
  }
  
  paginationHTML += `
    <button class="cmp-pagination-btn" onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>
      Next
    </button>
  `;
  
  container.innerHTML = paginationHTML;
}

function goToPage(page) {
  const totalPages = Math.ceil(filteredCmps.length / CMPS_PER_PAGE);
  if (page < 1 || page > totalPages) return;
  
  currentPage = page;
  displayCmps();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ==========================================
// Modal Functions
// ==========================================

function openCmpModal(cmpId = null) {
  const modal = document.getElementById('cmpModal');
  const title = document.getElementById('modalTitle');
  const form = document.getElementById('cmpForm');
  
  form.reset();
  clearFormErrors();
  
  if (cmpId) {
    const cmp = allCmps.find(c => c.CMP_ID === cmpId);
    if (!cmp) return;
    
    title.textContent = 'Edit CMP';
    document.getElementById('cmpId').value = cmp.CMP_ID;
    document.getElementById('cmpName').value = cmp.CMP_Name || '';
    document.getElementById('cmpCategory').value = cmp.CMP_Category || '';
    document.getElementById('cmpRegion').value = cmp.CMP_Region || '';
    document.getElementById('cmpTag').value = cmp.CMP_Tag || '';
    document.getElementById('cmpDescription').value = cmp.CMP_Description || '';
    document.getElementById('cmpInstructions').value = cmp.CMP_Instructions || '';
    document.getElementById('cmpOwner').value = cmp.CMP_Owner || '';
    document.getElementById('cmpProcessingTime').value = cmp.CMP_Processing_Time || '';
    document.getElementById('cmpPrerequisites').value = cmp.CMP_Prerequisites || '';
    document.getElementById('cmpIcon').value = cmp.CMP_Icon || '';
    document.getElementById('cmpTemplateUrl').value = cmp.CMP_Template_URL || '';
    document.getElementById('cmpAttachmentCode').value = cmp.CMP_Attachment_Code || '';
    document.getElementById('cmpAttachmentName').value = cmp.CMP_Attachment_Name || '';
    
    updateCharCounter('cmpDescription', 'descCharCount', 200);
    updateCharCounter('cmpInstructions', 'instructionsCharCount', 500);
  } else {
    title.textContent = 'Add New CMP';
  }
  
  modal.style.display = 'flex';
}

function closeCmpModal() {
  document.getElementById('cmpModal').style.display = 'none';
}

async function saveCmp() {
  if (!validateCmpForm()) return;
  
  const cmpId = document.getElementById('cmpId').value;
  const isEdit = !!cmpId;
  
  const cmpData = {
    CMP_ID: cmpId || generateId(),
    CMP_Name: document.getElementById('cmpName').value.trim(),
    CMP_Description: document.getElementById('cmpDescription').value.trim(),
    CMP_Category: document.getElementById('cmpCategory').value.trim(),
    CMP_Region: document.getElementById('cmpRegion').value.trim(),
    CMP_Tag: document.getElementById('cmpTag').value,
    CMP_Icon: document.getElementById('cmpIcon').value.trim(),
    CMP_Template_URL: document.getElementById('cmpTemplateUrl').value.trim(),
    CMP_Attachment_Code: document.getElementById('cmpAttachmentCode').value.trim(),
    CMP_Attachment_Name: document.getElementById('cmpAttachmentName').value.trim(),
    CMP_Instructions: document.getElementById('cmpInstructions').value.trim(),
    CMP_Owner: document.getElementById('cmpOwner').value.trim(),
    CMP_Processing_Time: document.getElementById('cmpProcessingTime').value.trim(),
    CMP_Prerequisites: document.getElementById('cmpPrerequisites').value.trim(),
    CMP_Added_By: isEdit ? (allCmps.find(c => c.CMP_ID === cmpId)?.CMP_Added_By || currentUserName) : currentUserName,
    CMP_Date_Added: isEdit ? (allCmps.find(c => c.CMP_ID === cmpId)?.CMP_Date_Added || new Date().toISOString()) : new Date().toISOString()
  };
  
  const saveBtn = document.getElementById('saveCmpBtn');
  saveBtn.disabled = true;
  saveBtn.textContent = 'Saving...';
  
  try {
    if (isEdit) {
      const index = allCmps.findIndex(c => c.CMP_ID === cmpId);
      if (index !== -1) {
        allCmps[index] = cmpData;
      }
    } else {
      allCmps.push(cmpData);
    }
    
    const headers = ['CMP_ID', 'CMP_Name', 'CMP_Description', 'CMP_Category', 'CMP_Region', 'CMP_Tag',
                     'CMP_Icon', 'CMP_Template_URL', 'CMP_Attachment_Code', 'CMP_Attachment_Name',
                     'CMP_Instructions', 'CMP_Owner', 'CMP_Processing_Time', 'CMP_Prerequisites',
                     'CMP_Added_By', 'CMP_Date_Added'];
    const csv = convertToCSV(allCmps, headers);
    
    await uploadCSVToConfluence(CSV_FILENAME_CMPS, csv);
    
    closeCmpModal();
    showToast('success', `CMP ${isEdit ? 'updated' : 'added'} successfully`);
    
    processCmps();
    applyFiltersAndDisplay();
    
  } catch (error) {
    console.error('Error saving CMP:', error);
    showToast('error', `Failed to ${isEdit ? 'update' : 'add'} CMP`);
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = 'Save CMP';
  }
}

function editCmp(cmpId) {
  openCmpModal(cmpId);
}

function deleteCmp(cmpId, cmpName) {
  cmpToDelete = cmpId;
  document.getElementById('deleteCmpName').textContent = cmpName;
  document.getElementById('deleteModal').style.display = 'flex';
}

function closeDeleteModal() {
  document.getElementById('deleteModal').style.display = 'none';
  cmpToDelete = null;
}

async function confirmDelete() {
  if (!cmpToDelete) return;
  
  const confirmBtn = document.getElementById('confirmDeleteBtn');
  confirmBtn.disabled = true;
  confirmBtn.textContent = 'Deleting...';
  
  try {
    allCmps = allCmps.filter(c => c.CMP_ID !== cmpToDelete);
    
    const headers = ['CMP_ID', 'CMP_Name', 'CMP_Description', 'CMP_Category', 'CMP_Region', 'CMP_Tag',
                     'CMP_Icon', 'CMP_Template_URL', 'CMP_Attachment_Code', 'CMP_Attachment_Name',
                     'CMP_Instructions', 'CMP_Owner', 'CMP_Processing_Time', 'CMP_Prerequisites',
                     'CMP_Added_By', 'CMP_Date_Added'];
    const csv = convertToCSV(allCmps, headers);
    
    await uploadCSVToConfluence(CSV_FILENAME_CMPS, csv);
    
    closeDeleteModal();
    showToast('success', 'CMP deleted successfully');
    
    processCmps();
    applyFiltersAndDisplay();
    
  } catch (error) {
    console.error('Error deleting CMP:', error);
    showToast('error', 'Failed to delete CMP');
  } finally {
    confirmBtn.disabled = false;
    confirmBtn.textContent = 'Delete CMP';
  }
}

// ==========================================
// Order & Attachment Modals
// ==========================================

function openOrderModal(cmpId) {
  const cmp = allCmps.find(c => c.CMP_ID === cmpId);
  if (!cmp || !cmp.CMP_Template_URL) return;
  
  // Track order click
  trackCmpAction(cmpId, cmp.CMP_Name, 'order_click', 'order');
  
  const modal = document.getElementById('orderModal');
  document.getElementById('orderModalTitle').textContent = `Order: ${cmp.CMP_Name}`;
  
  // Create iframe for template
  const iframe = document.createElement('iframe');
  iframe.src = cmp.CMP_Template_URL;
  iframe.style.width = '100%';
  iframe.style.height = '600px';
  iframe.style.border = 'none';
  
  const container = document.getElementById('orderIframeContainer');
  container.innerHTML = '';
  container.appendChild(iframe);
  
  modal.style.display = 'flex';
}

function closeOrderModal() {
  const container = document.getElementById('orderIframeContainer');
  container.innerHTML = ''; // Clear iframe
  document.getElementById('orderModal').style.display = 'none';
}

function openAttachmentModal(cmpId) {
  const cmp = allCmps.find(c => c.CMP_ID === cmpId);
  if (!cmp || !cmp.CMP_Attachment_Code) return;
  
  // Track attachment view
  trackCmpAction(cmpId, cmp.CMP_Name, 'view_attachment', 'attachment');
  
  const modal = document.getElementById('attachmentModal');
  const attachmentName = cmp.CMP_Attachment_Name || 'Attachment';
  document.getElementById('attachmentModalTitle').textContent = attachmentName;
  
  // Insert iframe embed code
  const container = document.getElementById('attachmentIframeContainer');
  container.innerHTML = cmp.CMP_Attachment_Code;
  
  modal.style.display = 'flex';
}

function closeAttachmentModal() {
  const container = document.getElementById('attachmentIframeContainer');
  container.innerHTML = ''; // Clear iframe
  document.getElementById('attachmentModal').style.display = 'none';
}

function copyTemplateLink(cmpId) {
  const cmp = allCmps.find(c => c.CMP_ID === cmpId);
  if (!cmp || !cmp.CMP_Template_URL) return;
  
  // Track copy action
  trackCmpAction(cmpId, cmp.CMP_Name, 'copy_link', 'copy');
  
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(cmp.CMP_Template_URL)
      .then(() => {
        showToast('success', `Template link copied for "${cmp.CMP_Name}"`);
      })
      .catch(err => {
        console.error('Failed to copy:', err);
        prompt('Copy this link:', cmp.CMP_Template_URL);
      });
  } else {
    prompt('Copy this link:', cmp.CMP_Template_URL);
  }
}

// ==========================================
// Usage Analytics Functions
// ==========================================

async function loadAnalytics() {
  try {
    console.log('Loading analytics...');
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_ANALYTICS);
    
    if (!csvData) {
      analyticsRecords = [];
    } else {
      analyticsRecords = parseCSV(csvData);
      console.log(`Loaded ${analyticsRecords.length} analytics records`);
    }
  } catch (error) {
    console.error('Error loading analytics:', error);
    analyticsRecords = [];
  }
}

async function trackCmpAction(cmpId, cmpName, action, environment) {
  try {
    const logEntry = {
      Log_ID: `track-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      Timestamp: new Date().toISOString(),
      CMP_ID: cmpId,
      CMP_Name: cmpName,
      User_Name: currentUserName,
      Action: action,
      Environment: environment
    };
    
    analyticsRecords.push(logEntry);
    
    saveAnalytics().catch(err => {
      console.error('Failed to save analytics:', err);
    });
    
  } catch (error) {
    console.error('Error tracking action:', error);
  }
}

async function saveAnalytics() {
  try {
    const ninetyDaysAgo = new Date();
    ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
    
    const filteredRecords = analyticsRecords.filter(r => {
      const recordDate = new Date(r.Timestamp);
      return recordDate > ninetyDaysAgo;
    });
    
    const headers = ['Log_ID', 'Timestamp', 'CMP_ID', 'CMP_Name', 'User_Name', 'Action', 'Environment'];
    const csv = convertToCSV(filteredRecords, headers);
    
    await uploadCSVToConfluence(CSV_FILENAME_ANALYTICS, csv);
    
    analyticsRecords = filteredRecords;
    
  } catch (error) {
    console.error('Error saving analytics:', error);
    throw error;
  }
}

// ==========================================
// Form Validation
// ==========================================

function validateCmpForm() {
  let isValid = true;
  
  const cmpName = document.getElementById('cmpName').value.trim();
  if (!cmpName) {
    showFieldError('nameError', 'CMP name is required');
    isValid = false;
  }
  
  const category = document.getElementById('cmpCategory').value.trim();
  if (!category) {
    showFieldError('categoryError', 'Please enter a category');
    isValid = false;
  }
  
  const region = document.getElementById('cmpRegion').value.trim();
  if (!region) {
    showFieldError('regionError', 'Please enter a region');
    isValid = false;
  }
  
  const description = document.getElementById('cmpDescription').value.trim();
  if (!description) {
    showFieldError('descriptionError', 'Description is required');
    isValid = false;
  }
  
  const iconUrl = document.getElementById('cmpIcon').value.trim();
  if (iconUrl && !isValidUrl(iconUrl)) {
    showFieldError('iconError', 'Please enter a valid URL');
    isValid = false;
  }
  
  const templateUrl = document.getElementById('cmpTemplateUrl').value.trim();
  if (templateUrl && !isValidUrl(templateUrl)) {
    showFieldError('templateError', 'Please enter a valid URL');
    isValid = false;
  }
  
  return isValid;
}

function isValidUrl(string) {
  try {
    new URL(string);
    return true;
  } catch (_) {
    return false;
  }
}

function showFieldError(errorId, message) {
  const errorEl = document.getElementById(errorId);
  errorEl.textContent = message;
  
  const inputId = errorId.replace('Error', '');
  const inputEl = document.getElementById(inputId);
  if (inputEl) {
    inputEl.classList.add('error');
  }
}

function clearFormErrors() {
  const errorElements = document.querySelectorAll('.cmp-error');
  errorElements.forEach(el => el.textContent = '');
  
  const inputElements = document.querySelectorAll('.cmp-input, .cmp-textarea');
  inputElements.forEach(el => el.classList.remove('error'));
}

function updateCharCounter(inputId, counterId, maxLength) {
  const input = document.getElementById(inputId);
  const counter = document.getElementById(counterId);
  counter.textContent = input.value.length;
}

// ==========================================
// Event Handlers
// ==========================================

function handleSearch() {
  searchQuery = document.getElementById('searchInput').value.trim();
  applyFiltersAndDisplay();
}

function handleSortChange(e) {
  currentSort = e.target.value;
  applyFiltersAndDisplay();
}

function setView(view) {
  currentView = view;
  
  document.getElementById('gridViewBtn').classList.toggle('active', view === 'grid');
  document.getElementById('listViewBtn').classList.toggle('active', view === 'list');
  
  displayCmps();
}

function toggleEditMode() {
  editMode = !editMode;
  deleteMode = false;
  updateModeButtons();
  displayCmps();
  showToast('info', editMode ? 'Edit mode enabled' : 'Edit mode disabled');
}

function toggleDeleteMode() {
  deleteMode = !deleteMode;
  editMode = false;
  updateModeButtons();
  displayCmps();
  showToast('info', deleteMode ? 'Delete mode enabled' : 'Delete mode disabled');
}

function updateModeButtons() {
  const editBtn = document.getElementById('editModeBtn');
  const deleteBtn = document.getElementById('deleteModeBtn');
  
  if (editBtn) {
    editBtn.classList.toggle('active', editMode);
    editBtn.querySelector('span').textContent = editMode ? 'Edit Mode: ON' : 'Edit Mode';
  }
  
  if (deleteBtn) {
    deleteBtn.classList.toggle('active', deleteMode);
    deleteBtn.querySelector('span').textContent = deleteMode ? 'Delete Mode: ON' : 'Delete Mode';
  }
}

async function handleRefresh() {
  const btn = document.getElementById('refreshBtn');
  btn.classList.add('loading');
  
  try {
    await loadAnalytics();
    await loadCmps();
    showToast('success', 'Data refreshed successfully');
  } catch (error) {
    console.error('Error refreshing data:', error);
    showToast('error', 'Failed to refresh data');
  } finally {
    btn.classList.remove('loading');
  }
}

function toggleConfigMenu() {
  const menu = document.getElementById('configMenu');
  menu.classList.toggle('active');
}

async function exportCmps() {
  try {
    const headers = ['CMP_ID', 'CMP_Name', 'CMP_Description', 'CMP_Category', 'CMP_Region', 'CMP_Tag',
                     'CMP_Icon', 'CMP_Template_URL', 'CMP_Attachment_Code', 'CMP_Attachment_Name',
                     'CMP_Instructions', 'CMP_Owner', 'CMP_Processing_Time', 'CMP_Prerequisites',
                     'CMP_Added_By', 'CMP_Date_Added'];
    const csv = convertToCSV(filteredCmps, headers);
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `cmp_library_export_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('success', 'CMPs exported successfully');
  } catch (error) {
    console.error('Error exporting CMPs:', error);
    showToast('error', 'Failed to export CMPs');
  }
}

async function exportAudit() {
  try {
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_AUDIT);
    
    if (!csvData) {
      showToast('warning', 'No audit log found');
      return;
    }
    
    const blob = new Blob([csvData], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `cmp_audit_log_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('success', 'Audit log exported successfully');
  } catch (error) {
    console.error('Error exporting audit log:', error);
    showToast('error', 'Failed to export audit log');
  }
}

function showSkeleton() {
  document.getElementById('skeletonContainer').style.display = 'grid';
  document.getElementById('cmpGrid').style.display = 'none';
  document.getElementById('cmpList').style.display = 'none';
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('pagination').style.display = 'none';
}

function hideSkeleton() {
  document.getElementById('skeletonContainer').style.display = 'none';
}

// ==========================================
// Auto-Refresh System
// ==========================================

function startAutoRefresh() {
  if (AUTO_REFRESH_INTERVAL > 0) {
    console.log(`Auto-refresh enabled: every ${AUTO_REFRESH_INTERVAL / 1000 / 60} minutes`);
    
    autoRefreshTimer = setInterval(async () => {
      console.log('Auto-refreshing data...');
      try {
        await loadAnalytics();
        await loadCmps();
        console.log('Auto-refresh complete');
      } catch (error) {
        console.error('Auto-refresh failed:', error);
      }
    }, AUTO_REFRESH_INTERVAL);
  }
}

function stopAutoRefresh() {
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer);
    autoRefreshTimer = null;
    console.log('Auto-refresh stopped');
  }
}

// ==========================================
// Setup Event Listeners
// ==========================================

function setupEventListeners() {
  document.getElementById('searchInput').addEventListener('input', debounce(handleSearch, SEARCH_DEBOUNCE_DELAY));
  document.getElementById('sortSelect').addEventListener('change', handleSortChange);
  
  document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
  document.getElementById('refreshBtn').addEventListener('click', handleRefresh);
  
  document.getElementById('gridViewBtn').addEventListener('click', () => setView('grid'));
  document.getElementById('listViewBtn').addEventListener('click', () => setView('list'));
  
  document.getElementById('configBtn').addEventListener('click', toggleConfigMenu);
  document.getElementById('addCmpBtn').addEventListener('click', () => {
    toggleConfigMenu();
    openCmpModal();
  });
  document.getElementById('editModeBtn').addEventListener('click', () => {
    toggleConfigMenu();
    toggleEditMode();
  });
  document.getElementById('deleteModeBtn').addEventListener('click', () => {
    toggleConfigMenu();
    toggleDeleteMode();
  });
  document.getElementById('exportCmpsBtn').addEventListener('click', () => {
    toggleConfigMenu();
    exportCmps();
  });
  document.getElementById('exportAuditBtn').addEventListener('click', () => {
    toggleConfigMenu();
    exportAudit();
  });
  
  document.getElementById('cmpDescription').addEventListener('input', () => {
    updateCharCounter('cmpDescription', 'descCharCount', 200);
  });
  document.getElementById('cmpInstructions').addEventListener('input', () => {
    updateCharCounter('cmpInstructions', 'instructionsCharCount', 500);
  });
  
  document.addEventListener('click', (e) => {
    const configDropdown = document.querySelector('.cmp-config-dropdown');
    if (!configDropdown.contains(e.target)) {
      document.getElementById('configMenu').classList.remove('active');
    }
  });
  
  document.getElementById('cmpModal').addEventListener('click', (e) => {
    if (e.target.id === 'cmpModal') {
      closeCmpModal();
    }
  });
  
  document.getElementById('orderModal').addEventListener('click', (e) => {
    if (e.target.id === 'orderModal') {
      closeOrderModal();
    }
  });
  
  document.getElementById('attachmentModal').addEventListener('click', (e) => {
    if (e.target.id === 'attachmentModal') {
      closeAttachmentModal();
    }
  });
  
  document.getElementById('deleteModal').addEventListener('click', (e) => {
    if (e.target.id === 'deleteModal') {
      closeDeleteModal();
    }
  });
}

// ==========================================
// Application Initialization
// ==========================================

async function initializeApp() {
  setupEventListeners();
  
  await loadCurrentUser();
  await loadAnalytics();
  await loadCmps();
  
  startAutoRefresh();
}

document.addEventListener('DOMContentLoaded', async () => {
  await initializeApp();
});

document.addEventListener('visibilitychange', () => {
  if (document.hidden) {
    stopAutoRefresh();
  } else {
    startAutoRefresh();
  }
});
