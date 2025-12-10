// ==========================================
// Configuration Constants
// ==========================================

// Confluence Configuration (UPDATE THESE)
const CONFLUENCE_BASE_URL = 'mysite/wiki';
const CONFLUENCE_PAGE_ID = '1235678990';
const CSV_FILENAME_LINKS = 'link_library.csv';
const CSV_FILENAME_AUDIT = 'link_audit_log.csv';

// Pagination
const LINKS_PER_PAGE = 12;

// Search debounce delay
const SEARCH_DEBOUNCE_DELAY = 300;

// Animation/Toast durations
const TOAST_DURATION = 5000;
const CARD_STAGGER_DELAY = 50;

// Category colors
const CATEGORY_COLORS = {
  'System': '#305edb',
  'Process': '#2a6b3c',
  'Tool': '#fab728',
  'Report': '#9a231a',
  'Other': '#64748b'
};

// Default categories
const DEFAULT_CATEGORIES = ['System', 'Process', 'Tool', 'Report', 'Other'];

// ==========================================
// Global State Variables
// ==========================================

let allLinks = [];
let filteredLinks = [];
let currentPage = 1;
let currentView = 'grid';
let currentSort = 'alphabetical';
let currentCategory = '';
let currentStatus = 'Active';
let searchQuery = '';
let editMode = false;
let deleteMode = false;
let currentUserName = '';
let dynamicCategories = [];
let searchDebounceTimer = null;
let linkToDelete = null;
let currentDetailsLinkId = null;
let currentDetailsLink = null;

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
  return 'link-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
}

function getCategoryColor(category) {
  return CATEGORY_COLORS[category] || CATEGORY_COLORS['Other'];
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
    'info': 'info'
  };
  
  const toast = document.createElement('div');
  toast.className = `ll-toast ll-toast-${type}`;
  toast.innerHTML = `
    <svg class="ll-toast-icon ll-icon"><use href="#icon-${iconMap[type]}"></use></svg>
    <div class="ll-toast-content">
      <h4 class="ll-toast-title">${titleMap[type]}</h4>
      <p class="ll-toast-message">${escapeHtml(message)}</p>
    </div>
    <button class="ll-toast-close" onclick="this.parentElement.remove()">
      <svg class="ll-icon ll-icon-sm"><use href="#icon-close"></use></svg>
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
    
    // Get attachments list
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
    
    // Download the file
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
    
    // Create FormData
    const formData = new FormData();
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8' });
    formData.append('file', blob, filename);
    formData.append('comment', `Updated via Link Library - ${new Date().toISOString()}`);
    
    console.log('Blob created. Size:', blob.size, 'bytes');
    
    // Check if file exists
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
      // Update existing attachment
      uploadUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment/${existingAttachment.id}/data`;
      console.log('Updating existing attachment. ID:', existingAttachment.id);
    } else {
      // Create new attachment
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
    
    // Verify the uploaded file
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

async function loadLinks() {
  try {
    console.log('=== LOADING LINKS ===');
    showSkeleton();
    
    // Fetch CSV from Confluence
    console.log('Fetching CSV from Confluence...');
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_LINKS);
    console.log('CSV data received. Length:', csvData.length);
    
    if (!csvData) {
      console.log('No CSV data found - first time setup');
      allLinks = [];
    } else {
      // Parse CSV to JSON
      console.log('Parsing CSV...');
      allLinks = parseCSV(csvData);
      console.log('Parsed links count:', allLinks.length);
      
      if (allLinks.length > 0) {
        console.log('Sample link:', {
          id: allLinks[0].Link_ID,
          name: allLinks[0].Link_App_Name,
          category: allLinks[0].Link_Category
        });
      }
    }
    
    // Process links
    console.log('Processing links...');
    processLinks();
    
    // Apply filters and display
    console.log('Applying filters and displaying...');
    applyFiltersAndDisplay();
    
    hideSkeleton();
    console.log('=== LOADING COMPLETE ===');
    
  } catch (error) {
    console.error('=== LOADING ERROR ===');
    console.error('Error loading links:', error);
    showToast('error', 'Failed to load links. Please check console for details.');
    hideSkeleton();
    
    // Show empty state
    document.getElementById('emptyState').style.display = 'block';
  }
}

function processLinks() {
  extractCategories();
}

function extractCategories() {
  const categoriesSet = new Set(DEFAULT_CATEGORIES);
  
  allLinks.forEach(link => {
    if (link.Link_Category) {
      categoriesSet.add(link.Link_Category);
    }
  });
  
  dynamicCategories = Array.from(categoriesSet).sort();
  populateCategoryDropdowns();
}

function populateCategoryDropdowns() {
  const filterDropdown = document.getElementById('categoryFilter');
  const formDropdown = document.getElementById('linkCategory');
  
  filterDropdown.innerHTML = '<option value="">All Categories</option>';
  formDropdown.innerHTML = '<option value="">Select Category</option>';
  
  dynamicCategories.forEach(category => {
    const filterOption = document.createElement('option');
    filterOption.value = category;
    filterOption.textContent = category;
    filterDropdown.appendChild(filterOption);
    
    const formOption = document.createElement('option');
    formOption.value = category;
    formOption.textContent = category;
    formDropdown.appendChild(formOption);
  });
}


// ==========================================
// Filtering & Sorting
// ==========================================

function applyFiltersAndDisplay() {
  filteredLinks = allLinks.filter(link => {
    if (currentStatus && link.Link_Status !== currentStatus) return false;
    
    if (currentCategory && link.Link_Category !== currentCategory) return false;
    
    if (searchQuery) {
      const query = searchQuery.toLowerCase();
      const matchesAppName = link.Link_App_Name?.toLowerCase().includes(query);
      const matchesDescription = link.Link_Description?.toLowerCase().includes(query);
      const matchesOwner = link.Link_Owner?.toLowerCase().includes(query);
      const matchesCategory = link.Link_Category?.toLowerCase().includes(query);
      
      if (!matchesAppName && !matchesDescription && !matchesOwner && !matchesCategory) {
        return false;
      }
    }
    
    return true;
  });
  
  sortLinks();
  updateResultsCount();
  updateClearFiltersButton();
  
  currentPage = 1;
  displayLinks();
}

function sortLinks() {
  switch (currentSort) {
    case 'alphabetical':
      filteredLinks.sort((a, b) => (a.Link_App_Name || '').localeCompare(b.Link_App_Name || ''));
      break;
    case 'alphabetical-desc':
      filteredLinks.sort((a, b) => (b.Link_App_Name || '').localeCompare(a.Link_App_Name || ''));
      break;
    case 'newest':
      filteredLinks.sort((a, b) => new Date(b.Link_Date_Added || 0) - new Date(a.Link_Date_Added || 0));
      break;
    case 'oldest':
      filteredLinks.sort((a, b) => new Date(a.Link_Date_Added || 0) - new Date(b.Link_Date_Added || 0));
      break;
    case 'owner':
      filteredLinks.sort((a, b) => (a.Link_Owner || '').localeCompare(b.Link_Owner || ''));
      break;
  }
}

function updateResultsCount() {
  const resultsEl = document.getElementById('resultsCount');
  const total = filteredLinks.length;
  const start = (currentPage - 1) * LINKS_PER_PAGE + 1;
  const end = Math.min(currentPage * LINKS_PER_PAGE, total);
  
  if (total === 0) {
    resultsEl.textContent = 'No links found';
  } else {
    resultsEl.textContent = `Showing ${start}-${end} of ${total} link${total !== 1 ? 's' : ''}`;
  }
}

function updateClearFiltersButton() {
  const btn = document.getElementById('clearFiltersBtn');
  const hasFilters = searchQuery || currentCategory || (currentStatus !== 'Active');
  btn.style.display = hasFilters ? 'block' : 'none';
}

function clearFilters() {
  searchQuery = '';
  currentCategory = '';
  currentStatus = 'Active';
  
  document.getElementById('searchInput').value = '';
  document.getElementById('categoryFilter').value = '';
  document.getElementById('statusFilter').value = 'Active';
  
  applyFiltersAndDisplay();
}

// ==========================================
// Display Functions
// ==========================================

function displayLinks() {
  if (currentView === 'grid') {
    displayGridView();
  } else {
    displayListView();
  }
  
  updatePagination();
}

function displayGridView() {
  const container = document.getElementById('linkGrid');
  const emptyState = document.getElementById('emptyState');
  
  container.style.display = 'grid';
  document.getElementById('linkList').style.display = 'none';
  
  if (filteredLinks.length === 0) {
    container.style.display = 'none';
    emptyState.style.display = 'block';
    return;
  }
  
  emptyState.style.display = 'none';
  
  const start = (currentPage - 1) * LINKS_PER_PAGE;
  const end = start + LINKS_PER_PAGE;
  const linksToShow = filteredLinks.slice(start, end);
  
  container.innerHTML = linksToShow.map((link, index) => createLinkCard(link, index)).join('');
}

function createLinkCard(link, index) {
  const categoryColor = getCategoryColor(link.Link_Category);
  const hasProd = link.Link_Prod && link.Link_Prod.trim();
  const hasUat = link.Link_UAT && link.Link_UAT.trim();
  
  const iconHtml = link.Link_Avatar 
    ? `<img src="${escapeHtml(link.Link_Avatar)}" alt="${escapeHtml(link.Link_App_Name)}" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" />
       <svg class="ll-icon" style="display: none; color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`
    : `<svg class="ll-icon" style="color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`;
  
  return `
    <div class="ll-link-card" style="animation-delay: ${index * CARD_STAGGER_DELAY}ms">
      <div class="ll-card-header">
        <div class="ll-app-icon" style="background: ${categoryColor}20;">
          ${iconHtml}
        </div>
        <div class="ll-card-title-section">
          <div class="ll-card-title-row">
            <h3 class="ll-card-title">${escapeHtml(link.Link_App_Name)}</h3>
            <span class="ll-badge-category" style="background: ${categoryColor}; color: white;">
              ${escapeHtml(link.Link_Category)}
            </span>
          </div>
        </div>
      </div>
      
      <p class="ll-card-description">${escapeHtml(link.Link_Description || 'No description available')}</p>
      
      <div class="ll-card-meta">
        ${link.Link_Owner ? `
          <div class="ll-card-meta-item">
            <svg class="ll-icon ll-icon-xs"><use href="#icon-user"></use></svg>
            <span>${escapeHtml(link.Link_Owner)}</span>
          </div>
        ` : ''}
        ${link.Link_Date_Added ? `
          <div class="ll-card-meta-item">
            <svg class="ll-icon ll-icon-xs"><use href="#icon-calendar"></use></svg>
            <span>${formatDate(link.Link_Date_Added)}</span>
          </div>
        ` : ''}
      </div>
      
      <div class="ll-card-actions">
        ${hasProd 
          ? `<button class="ll-btn ll-btn-primary ll-btn-sm" onclick="openLinkUrl('${link.Link_ID}', 'prod'); event.stopPropagation();">
               <svg class="ll-icon ll-icon-xs"><use href="#icon-rocket"></use></svg>
               PROD
             </button>`
          : `<button class="ll-btn ll-btn-na ll-btn-sm" disabled>
               PROD (N/A)
             </button>`
        }
        
        ${hasUat 
          ? `<button class="ll-btn ll-btn-gray ll-btn-sm" onclick="openLinkUrl('${link.Link_ID}', 'uat'); event.stopPropagation();">
               <svg class="ll-icon ll-icon-xs"><use href="#icon-flask"></use></svg>
               UAT
             </button>`
          : `<button class="ll-btn ll-btn-na ll-btn-sm" disabled>
               UAT (N/A)
             </button>`
        }
        
        <button class="ll-icon-btn" onclick="openDetailsModal('${link.Link_ID}'); event.stopPropagation();" title="View details">
          <svg class="ll-icon ll-icon-sm"><use href="#icon-info"></use></svg>
        </button>
        
        ${editMode || deleteMode ? `
          <div class="ll-card-edit-actions">
            ${editMode ? `
              <button class="ll-icon-btn" onclick="editLink('${link.Link_ID}'); event.stopPropagation();" title="Edit">
                <svg class="ll-icon ll-icon-sm"><use href="#icon-edit"></use></svg>
              </button>
            ` : ''}
            ${deleteMode ? `
              <button class="ll-icon-btn danger" onclick="deleteLink('${link.Link_ID}', '${escapeHtml(link.Link_App_Name)}'); event.stopPropagation();" title="Delete">
                <svg class="ll-icon ll-icon-sm"><use href="#icon-trash"></use></svg>
              </button>
            ` : ''}
          </div>
        ` : ''}
      </div>
    </div>
  `;
}

function displayListView() {
  const container = document.getElementById('linkList');
  const emptyState = document.getElementById('emptyState');
  
  container.style.display = 'block';
  document.getElementById('linkGrid').style.display = 'none';
  
  if (filteredLinks.length === 0) {
    container.style.display = 'none';
    emptyState.style.display = 'block';
    return;
  }
  
  emptyState.style.display = 'none';
  
  const start = (currentPage - 1) * LINKS_PER_PAGE;
  const end = start + LINKS_PER_PAGE;
  const linksToShow = filteredLinks.slice(start, end);
  
  container.innerHTML = linksToShow.map(link => createListItem(link)).join('');
}

function createListItem(link) {
  const categoryColor = getCategoryColor(link.Link_Category);
  const hasProd = link.Link_Prod && link.Link_Prod.trim();
  const hasUat = link.Link_UAT && link.Link_UAT.trim();
  
  const iconHtml = link.Link_Avatar 
    ? `<img src="${escapeHtml(link.Link_Avatar)}" alt="${escapeHtml(link.Link_App_Name)}" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" />
       <svg class="ll-icon" style="display: none; color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`
    : `<svg class="ll-icon" style="color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`;
  
  return `
    <div class="ll-list-item">
      <div class="ll-list-icon" style="background: ${categoryColor}20;">
        ${iconHtml}
      </div>
      
      <div class="ll-list-content">
        <h4 class="ll-list-title">${escapeHtml(link.Link_App_Name)}</h4>
        <p class="ll-list-description">${escapeHtml(link.Link_Description || 'No description')}</p>
      </div>
      
      <div class="ll-list-owner">${escapeHtml(link.Link_Owner || '-')}</div>
      
      <span class="ll-badge-category" style="background: ${categoryColor}; color: white;">
        ${escapeHtml(link.Link_Category)}
      </span>
      
      <div class="ll-list-date">${formatDate(link.Link_Date_Added)}</div>
      
      <div class="ll-list-actions">
        ${hasProd 
          ? `<button class="ll-btn ll-btn-primary ll-btn-sm" onclick="openLinkUrl('${link.Link_ID}', 'prod')">
               <svg class="ll-icon ll-icon-xs"><use href="#icon-rocket"></use></svg>
             </button>`
          : ''
        }
        ${hasUat 
          ? `<button class="ll-btn ll-btn-gray ll-btn-sm" onclick="openLinkUrl('${link.Link_ID}', 'uat')">
               <svg class="ll-icon ll-icon-xs"><use href="#icon-flask"></use></svg>
             </button>`
          : ''
        }
        <button class="ll-icon-btn" onclick="openDetailsModal('${link.Link_ID}')" title="Details">
          <svg class="ll-icon ll-icon-sm"><use href="#icon-info"></use></svg>
        </button>
        ${editMode ? `
          <button class="ll-icon-btn" onclick="editLink('${link.Link_ID}')" title="Edit">
            <svg class="ll-icon ll-icon-sm"><use href="#icon-edit"></use></svg>
          </button>
        ` : ''}
        ${deleteMode ? `
          <button class="ll-icon-btn danger" onclick="deleteLink('${link.Link_ID}', '${escapeHtml(link.Link_App_Name)}')" title="Delete">
            <svg class="ll-icon ll-icon-sm"><use href="#icon-trash"></use></svg>
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
  const totalPages = Math.ceil(filteredLinks.length / LINKS_PER_PAGE);
  
  if (totalPages <= 1) {
    container.style.display = 'none';
    return;
  }
  
  container.style.display = 'flex';
  
  let paginationHTML = `
    <button class="ll-pagination-btn" onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>
      Previous
    </button>
  `;
  
  for (let i = 1; i <= Math.min(totalPages, 5); i++) {
    paginationHTML += `
      <button class="ll-page-btn ${i === currentPage ? 'active' : ''}" onclick="goToPage(${i})">
        ${i}
      </button>
    `;
  }
  
  if (totalPages > 5) {
    paginationHTML += '<span style="padding: 8px;">...</span>';
    paginationHTML += `
      <button class="ll-page-btn ${totalPages === currentPage ? 'active' : ''}" onclick="goToPage(${totalPages})">
        ${totalPages}
      </button>
    `;
  }
  
  paginationHTML += `
    <button class="ll-pagination-btn" onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>
      Next
    </button>
  `;
  
  container.innerHTML = paginationHTML;
}

function goToPage(page) {
  const totalPages = Math.ceil(filteredLinks.length / LINKS_PER_PAGE);
  if (page < 1 || page > totalPages) return;
  
  currentPage = page;
  displayLinks();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ==========================================
// Modal Functions
// ==========================================

function openLinkModal(linkId = null) {
  const modal = document.getElementById('linkModal');
  const title = document.getElementById('modalTitle');
  const form = document.getElementById('linkForm');
  
  form.reset();
  clearFormErrors();
  
  if (linkId) {
    const link = allLinks.find(l => l.Link_ID === linkId);
    if (!link) return;
    
    title.textContent = 'Edit Link';
    document.getElementById('linkId').value = link.Link_ID;
    document.getElementById('linkAppName').value = link.Link_App_Name || '';
    document.getElementById('linkCategory').value = link.Link_Category || '';
    document.getElementById('linkDescription').value = link.Link_Description || '';
    document.getElementById('linkDetails').value = link.Link_Details || '';
    document.getElementById('linkOwner').value = link.Link_Owner || '';
    document.getElementById('linkAvatar').value = link.Link_Avatar || '';
    document.getElementById('linkProd').value = link.Link_Prod || '';
    document.getElementById('linkUat').value = link.Link_UAT || '';
    
    updateCharCounter('linkDescription', 'descCharCount', 200);
    updateCharCounter('linkDetails', 'detailsCharCount', 1000);
  } else {
    title.textContent = 'Add New Link';
    document.getElementById('linkOwner').value = currentUserName;
  }
  
  modal.style.display = 'flex';
}

function closeLinkModal() {
  document.getElementById('linkModal').style.display = 'none';
}

async function saveLink() {
  if (!validateLinkForm()) return;
  
  const linkId = document.getElementById('linkId').value;
  const isEdit = !!linkId;
  
  const linkData = {
    Link_ID: linkId || generateId(),
    Link_App_Name: document.getElementById('linkAppName').value.trim(),
    Link_Description: document.getElementById('linkDescription').value.trim(),
    Link_Category: document.getElementById('linkCategory').value,
    Link_Avatar: document.getElementById('linkAvatar').value.trim(),
    Link_Prod: document.getElementById('linkProd').value.trim(),
    Link_UAT: document.getElementById('linkUat').value.trim(),
    Link_Details: document.getElementById('linkDetails').value.trim(),
    Link_Owner: document.getElementById('linkOwner').value.trim(),
    Link_Added_By: isEdit ? (allLinks.find(l => l.Link_ID === linkId)?.Link_Added_By || currentUserName) : currentUserName,
    Link_Date_Added: isEdit ? (allLinks.find(l => l.Link_ID === linkId)?.Link_Date_Added || new Date().toISOString()) : new Date().toISOString(),
    Link_Status: 'Active',
    Link_Access_Level: 'Public'
  };
  
  const saveBtn = document.getElementById('saveLinkBtn');
  saveBtn.disabled = true;
  saveBtn.textContent = 'Saving...';
  
  try {
    if (isEdit) {
      const index = allLinks.findIndex(l => l.Link_ID === linkId);
      if (index !== -1) {
        allLinks[index] = linkData;
      }
    } else {
      allLinks.push(linkData);
    }
    
    const headers = ['Link_ID', 'Link_App_Name', 'Link_Description', 'Link_Category', 'Link_Avatar', 
                     'Link_Prod', 'Link_UAT', 'Link_Details', 'Link_Owner', 'Link_Added_By', 
                     'Link_Date_Added', 'Link_Status', 'Link_Access_Level'];
    const csv = convertToCSV(allLinks, headers);
    
    await uploadCSVToConfluence(CSV_FILENAME_LINKS, csv);
    
    closeLinkModal();
    showToast('success', `Link ${isEdit ? 'updated' : 'added'} successfully`);
    
    processLinks();
    applyFiltersAndDisplay();
    
  } catch (error) {
    console.error('Error saving link:', error);
    showToast('error', `Failed to ${isEdit ? 'update' : 'add'} link`);
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = 'Save Link';
  }
}

function editLink(linkId) {
  openLinkModal(linkId);
}

function deleteLink(linkId, linkName) {
  linkToDelete = linkId;
  document.getElementById('deleteLinkName').textContent = linkName;
  document.getElementById('deleteModal').style.display = 'flex';
}

function closeDeleteModal() {
  document.getElementById('deleteModal').style.display = 'none';
  linkToDelete = null;
}

async function confirmDelete() {
  if (!linkToDelete) return;
  
  const confirmBtn = document.getElementById('confirmDeleteBtn');
  confirmBtn.disabled = true;
  confirmBtn.textContent = 'Deleting...';
  
  try {
    allLinks = allLinks.filter(l => l.Link_ID !== linkToDelete);
    
    const headers = ['Link_ID', 'Link_App_Name', 'Link_Description', 'Link_Category', 'Link_Avatar', 
                     'Link_Prod', 'Link_UAT', 'Link_Details', 'Link_Owner', 'Link_Added_By', 
                     'Link_Date_Added', 'Link_Status', 'Link_Access_Level'];
    const csv = convertToCSV(allLinks, headers);
    
    await uploadCSVToConfluence(CSV_FILENAME_LINKS, csv);
    
    closeDeleteModal();
    showToast('success', 'Link deleted successfully');
    
    processLinks();
    applyFiltersAndDisplay();
    
  } catch (error) {
    console.error('Error deleting link:', error);
    showToast('error', 'Failed to delete link');
  } finally {
    confirmBtn.disabled = false;
    confirmBtn.textContent = 'Delete Link';
  }
}

// ==========================================
// Details Modal
// ==========================================

function openDetailsModal(linkId) {
  const link = allLinks.find(l => l.Link_ID === linkId);
  if (!link) return;
  
  currentDetailsLinkId = linkId;
  currentDetailsLink = link;
  
  const modal = document.getElementById('detailsModal');
  const categoryColor = getCategoryColor(link.Link_Category);
  
  const iconHtml = link.Link_Avatar 
    ? `<img src="${escapeHtml(link.Link_Avatar)}" alt="${escapeHtml(link.Link_App_Name)}" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" />
       <svg class="ll-icon" style="display: none; color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`
    : `<svg class="ll-icon" style="color: ${categoryColor};"><use href="#icon-app-default"></use></svg>`;
  
  document.getElementById('detailsIcon').innerHTML = iconHtml;
  document.getElementById('detailsIcon').style.background = `${categoryColor}20`;
  document.getElementById('detailsAppName').textContent = link.Link_App_Name;
  document.getElementById('detailsDescription').textContent = link.Link_Description || 'No description available';
  
  const additionalSection = document.getElementById('detailsAdditionalSection');
  if (link.Link_Details) {
    document.getElementById('detailsAdditional').textContent = link.Link_Details;
    additionalSection.style.display = 'block';
  } else {
    additionalSection.style.display = 'none';
  }
  
  document.getElementById('detailsOwner').textContent = link.Link_Owner || '-';
  document.getElementById('detailsDate').textContent = formatDate(link.Link_Date_Added);
  document.getElementById('detailsCategory').textContent = link.Link_Category;
  
  const prodContainer = document.getElementById('prodUrlContainer');
  const uatContainer = document.getElementById('uatUrlContainer');
  const openProdBtn = document.getElementById('openProdBtn');
  const openUatBtn = document.getElementById('openUatBtn');
  
  if (link.Link_Prod && link.Link_Prod.trim()) {
    document.getElementById('prodUrlText').textContent = link.Link_Prod;
    prodContainer.style.display = 'flex';
    openProdBtn.style.display = 'inline-flex';
  } else {
    prodContainer.style.display = 'none';
    openProdBtn.style.display = 'none';
  }
  
  if (link.Link_UAT && link.Link_UAT.trim()) {
    document.getElementById('uatUrlText').textContent = link.Link_UAT;
    uatContainer.style.display = 'flex';
    openUatBtn.style.display = 'inline-flex';
  } else {
    uatContainer.style.display = 'none';
    openUatBtn.style.display = 'none';
  }
  
  modal.style.display = 'flex';
}

function closeDetailsModal() {
  document.getElementById('detailsModal').style.display = 'none';
  currentDetailsLinkId = null;
  currentDetailsLink = null;
}

function copyUrl(type) {
  const url = type === 'prod' ? currentDetailsLink.Link_Prod : currentDetailsLink.Link_UAT;
  
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(url)
      .then(() => {
        showToast('success', `${type.toUpperCase()} URL copied to clipboard`);
      })
      .catch(err => {
        console.error('Failed to copy:', err);
        prompt('Copy this URL:', url);
      });
  } else {
    prompt('Copy this URL:', url);
  }
}

function openUrl(type) {
  const url = type === 'prod' ? currentDetailsLink.Link_Prod : currentDetailsLink.Link_UAT;
  if (url) {
    window.open(url, '_blank', 'noopener,noreferrer');
  }
}

// ==========================================
// Link Actions
// ==========================================

function openLinkUrl(linkId, type) {
  const link = allLinks.find(l => l.Link_ID === linkId);
  if (!link) return;
  
  const url = type === 'prod' ? link.Link_Prod : link.Link_UAT;
  if (url) {
    window.open(url, '_blank', 'noopener,noreferrer');
  }
}

function shareLink(linkId) {
  const link = allLinks.find(l => l.Link_ID === linkId);
  if (!link) return;
  
  const baseUrl = window.location.href.split('?')[0];
  const shareUrl = `${baseUrl}?link=${linkId}`;
  
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(shareUrl)
      .then(() => {
        showToast('success', `Share link copied! Anyone with this link can access "${link.Link_App_Name}" details`);
      })
      .catch(err => {
        console.error('Failed to copy to clipboard:', err);
        prompt('Copy this link to share:', shareUrl);
      });
  } else {
    prompt('Copy this link to share:', shareUrl);
  }
}

// ==========================================
// Form Validation
// ==========================================

function validateLinkForm() {
  let isValid = true;
  
  const appName = document.getElementById('linkAppName').value.trim();
  if (!appName) {
    showFieldError('appNameError', 'Application name is required');
    isValid = false;
  }
  
  const category = document.getElementById('linkCategory').value;
  if (!category) {
    showFieldError('categoryError', 'Please select a category');
    isValid = false;
  }
  
  const description = document.getElementById('linkDescription').value.trim();
  if (!description) {
    showFieldError('descriptionError', 'Description is required');
    isValid = false;
  }
  
  const avatarUrl = document.getElementById('linkAvatar').value.trim();
  if (avatarUrl && !isValidUrl(avatarUrl)) {
    showFieldError('avatarError', 'Please enter a valid URL');
    isValid = false;
  }
  
  const prodUrl = document.getElementById('linkProd').value.trim();
  if (prodUrl && !isValidUrl(prodUrl)) {
    showFieldError('prodError', 'Please enter a valid URL');
    isValid = false;
  }
  
  const uatUrl = document.getElementById('linkUat').value.trim();
  if (uatUrl && !isValidUrl(uatUrl)) {
    showFieldError('uatError', 'Please enter a valid URL');
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
  const errorElements = document.querySelectorAll('.ll-error');
  errorElements.forEach(el => el.textContent = '');
  
  const inputElements = document.querySelectorAll('.ll-input, .ll-textarea');
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

function handleCategoryFilter(e) {
  currentCategory = e.target.value;
  applyFiltersAndDisplay();
}

function handleStatusFilter(e) {
  currentStatus = e.target.value;
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
  
  displayLinks();
}

function toggleEditMode() {
  editMode = !editMode;
  deleteMode = false;
  updateModeButtons();
  displayLinks();
  showToast('info', editMode ? 'Edit mode enabled' : 'Edit mode disabled');
}

function toggleDeleteMode() {
  deleteMode = !deleteMode;
  editMode = false;
  updateModeButtons();
  displayLinks();
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
    await loadLinks();
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

async function exportLinks() {
  try {
    const headers = ['Link_ID', 'Link_App_Name', 'Link_Description', 'Link_Category', 'Link_Avatar', 
                     'Link_Prod', 'Link_UAT', 'Link_Details', 'Link_Owner', 'Link_Added_By', 
                     'Link_Date_Added', 'Link_Status', 'Link_Access_Level'];
    const csv = convertToCSV(filteredLinks, headers);
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `link_library_export_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('success', 'Links exported successfully');
  } catch (error) {
    console.error('Error exporting links:', error);
    showToast('error', 'Failed to export links');
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
    a.download = `link_audit_log_${new Date().toISOString().split('T')[0]}.csv`;
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
  document.getElementById('linkGrid').style.display = 'none';
  document.getElementById('linkList').style.display = 'none';
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('pagination').style.display = 'none';
}

function hideSkeleton() {
  document.getElementById('skeletonContainer').style.display = 'none';
}

// ==========================================
// Setup Event Listeners
// ==========================================

function setupEventListeners() {
  document.getElementById('searchInput').addEventListener('input', debounce(handleSearch, SEARCH_DEBOUNCE_DELAY));
  document.getElementById('categoryFilter').addEventListener('change', handleCategoryFilter);
  document.getElementById('statusFilter').addEventListener('change', handleStatusFilter);
  document.getElementById('sortSelect').addEventListener('change', handleSortChange);
  
  document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
  document.getElementById('refreshBtn').addEventListener('click', handleRefresh);
  
  document.getElementById('gridViewBtn').addEventListener('click', () => setView('grid'));
  document.getElementById('listViewBtn').addEventListener('click', () => setView('list'));
  
  document.getElementById('configBtn').addEventListener('click', toggleConfigMenu);
  document.getElementById('addLinkBtn').addEventListener('click', () => {
    toggleConfigMenu();
    openLinkModal();
  });
  document.getElementById('editModeBtn').addEventListener('click', () => {
    toggleConfigMenu();
    toggleEditMode();
  });
  document.getElementById('deleteModeBtn').addEventListener('click', () => {
    toggleConfigMenu();
    toggleDeleteMode();
  });
  document.getElementById('exportLinksBtn').addEventListener('click', () => {
    toggleConfigMenu();
    exportLinks();
  });
  document.getElementById('exportAuditBtn').addEventListener('click', () => {
    toggleConfigMenu();
    exportAudit();
  });
  
  document.getElementById('linkDescription').addEventListener('input', () => {
    updateCharCounter('linkDescription', 'descCharCount', 200);
  });
  document.getElementById('linkDetails').addEventListener('input', () => {
    updateCharCounter('linkDetails', 'detailsCharCount', 1000);
  });
  
  document.addEventListener('click', (e) => {
    const configDropdown = document.querySelector('.ll-config-dropdown');
    if (!configDropdown.contains(e.target)) {
      document.getElementById('configMenu').classList.remove('active');
    }
  });
  
  document.getElementById('linkModal').addEventListener('click', (e) => {
    if (e.target.id === 'linkModal') {
      closeLinkModal();
    }
  });
  
  document.getElementById('detailsModal').addEventListener('click', (e) => {
    if (e.target.id === 'detailsModal') {
      closeDetailsModal();
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
  populateCategoryDropdowns();
  
  await loadCurrentUser();
  await loadLinks();
  
  const urlParams = new URLSearchParams(window.location.search);
  const linkId = urlParams.get('link');
  if (linkId) {
    setTimeout(() => {
      const link = allLinks.find(l => l.Link_ID === linkId);
      if (link) {
        openDetailsModal(linkId);
      }
    }, 500);
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  await initializeApp();
});
