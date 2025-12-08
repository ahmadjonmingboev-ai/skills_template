/* ==========================================
   Video Library - Main JavaScript
   ========================================== */

// Configuration Constants
const VIDEOS_PER_PAGE = 12;
const ATTESTATION_WARNING_DAYS = 30;
const ATTESTATION_CRITICAL_DAYS = 0;
const SEARCH_DEBOUNCE_DELAY = 300;
const MODAL_ANIMATION_DURATION = 300;
const TOAST_DURATION = 5000;
const CARD_STAGGER_DELAY = 50;

const STORAGE_VIEW_MODE = 'videoLibrary_viewMode';
const STORAGE_VIDEOS_PER_PAGE = 'videoLibrary_perPage';

const VIDEO_CATEGORIES = [
  'Process',
  'Policy',
  'Product',
  'Technical',
  'Compliance',
  'Other'
];

const CONFLUENCE_BASE_URL = 'mysite/wiki';
const CONFLUENCE_PAGE_ID = '1235678990';
const CSV_FILENAME_VIDEOS = 'training_videos.csv';
const CSV_FILENAME_AUDIT = 'video_audit_log.csv';

// State Management
let allVideos = [];
let filteredVideos = [];
let currentPage = 1;
let currentView = 'grid';
let currentSort = 'newest';
let currentCategory = '';
let searchQuery = '';
let editMode = false;
let deleteMode = false;
let searchTimeout = null;

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  initializeApp();
});

function initializeApp() {
  // Load saved preferences
  loadPreferences();
  
  // Setup event listeners
  setupEventListeners();
  
  // Populate category dropdown
  populateCategoryDropdown();
  
  // Load videos
  loadVideos();
}

// ==========================================
// Event Listeners
// ==========================================

function setupEventListeners() {
  // Search
  const searchInput = document.getElementById('searchInput');
  const searchClear = document.getElementById('searchClear');
  searchInput.addEventListener('input', handleSearch);
  searchClear.addEventListener('click', clearSearch);
  
  // View toggle
  document.getElementById('gridViewBtn').addEventListener('click', () => setView('grid'));
  document.getElementById('listViewBtn').addEventListener('click', () => setView('list'));
  
  // Config menu
  const configBtn = document.getElementById('configBtn');
  const configMenu = document.getElementById('configMenu');
  configBtn.addEventListener('click', () => toggleDropdown(configMenu));
  
  // Close dropdown when clicking outside
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.vl-config-dropdown')) {
      configMenu.classList.remove('show');
    }
  });
  
  // Config menu items
  document.getElementById('addVideoBtn').addEventListener('click', () => openVideoModal());
  document.getElementById('editModeBtn').addEventListener('click', toggleEditMode);
  document.getElementById('deleteModeBtn').addEventListener('click', toggleDeleteMode);
  document.getElementById('exportVideosBtn').addEventListener('click', exportVideos);
  document.getElementById('exportAuditBtn').addEventListener('click', exportAuditLog);
  document.getElementById('toggleAttestationDashboard').addEventListener('click', toggleAttestationDashboardVisibility);
  
  // Filters
  document.getElementById('categoryFilter').addEventListener('change', handleCategoryFilter);
  document.getElementById('sortSelect').addEventListener('change', handleSort);
  document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
  
  // Attestation widget toggle
  document.getElementById('widgetToggle').addEventListener('click', toggleAttestationWidget);
  
  // Pagination
  document.getElementById('prevPageBtn').addEventListener('click', () => changePage(currentPage - 1));
  document.getElementById('nextPageBtn').addEventListener('click', () => changePage(currentPage + 1));
  
  // Video modal
  document.getElementById('modalClose').addEventListener('click', closeVideoModal);
  document.getElementById('modalCancel').addEventListener('click', closeVideoModal);
  document.getElementById('videoForm').addEventListener('submit', handleVideoSubmit);
  
  // Form field listeners
  setupFormListeners();
  
  // Delete modal
  document.getElementById('deleteModalClose').addEventListener('click', closeDeleteModal);
  document.getElementById('deleteCancel').addEventListener('click', closeDeleteModal);
  document.getElementById('deleteConfirm').addEventListener('click', confirmDelete);
  
  // Empty state action
  document.getElementById('emptyActionBtn').addEventListener('click', () => openVideoModal());
  
  // Close modals on overlay click
  document.getElementById('videoModal').addEventListener('click', (e) => {
    if (e.target.id === 'videoModal') closeVideoModal();
  });
  document.getElementById('deleteModal').addEventListener('click', (e) => {
    if (e.target.id === 'deleteModal') closeDeleteModal();
  });
  
  // ESC key to close modals
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeVideoModal();
      closeDeleteModal();
    }
  });
}

function setupFormListeners() {
  // Character counters
  const titleInput = document.getElementById('videoTitle');
  const descInput = document.getElementById('videoDescription');
  
  titleInput.addEventListener('input', () => {
    document.getElementById('titleCounter').textContent = titleInput.value.length;
    validateField(titleInput, 'titleError', 'Title is required');
  });
  
  descInput.addEventListener('input', () => {
    document.getElementById('descCounter').textContent = descInput.value.length;
    validateField(descInput, 'descError', 'Description is required');
  });
  
  // Validation on blur
  document.getElementById('videoAuthor').addEventListener('blur', (e) => {
    validateField(e.target, 'authorError', 'Author name is required');
  });
  
  document.getElementById('videoCategory').addEventListener('change', (e) => {
    validateField(e.target, 'categoryError', 'Please select a category');
  });
  
  document.getElementById('videoEmbed').addEventListener('blur', (e) => {
    validateEmbedCode(e.target);
  });
  
  document.getElementById('videoDuration').addEventListener('blur', (e) => {
    validateDuration(e.target);
  });
  
  document.getElementById('videoDocument').addEventListener('blur', (e) => {
    validateURL(e.target, 'documentError');
  });
  
  document.getElementById('videoThumbnail').addEventListener('blur', (e) => {
    validateURL(e.target, 'thumbnailError');
  });
}

// ==========================================
// Data Loading & API Integration
// ==========================================

async function loadVideos() {
  try {
    showSkeleton();
    
    // Fetch CSV from Confluence
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_VIDEOS);
    
    // Parse CSV to JSON
    allVideos = parseCSV(csvData);
    
    // Process videos (calculate attestation, etc.)
    processVideos();
    
    // Apply filters and display
    applyFiltersAndDisplay();
    
    hideSkeleton();
  } catch (error) {
    console.error('Error loading videos:', error);
    hideSkeleton();
    showToast('error', 'Failed to load videos. Please refresh the page.');
    showEmptyState('error');
  }
}

async function fetchCSVFromConfluence(filename) {
  try {
    // Get attachments from page
    const attachmentsUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment`;
    const response = await fetch(attachmentsUrl);
    
    if (!response.ok) {
      throw new Error('Failed to fetch attachments');
    }
    
    const data = await response.json();
    const attachment = data.results.find(att => att.title === filename);
    
    if (!attachment) {
      // If file doesn't exist, return empty (first time use)
      return '';
    }
    
    // Download file content
    const downloadUrl = attachment._links.download;
    const fileResponse = await fetch(downloadUrl);
    
    if (!fileResponse.ok) {
      throw new Error('Failed to download file');
    }
    
    return await fileResponse.text();
  } catch (error) {
    console.error('Error fetching CSV:', error);
    throw error;
  }
}

function parseCSV(csvText) {
  if (!csvText || csvText.trim() === '') {
    return [];
  }
  
  const lines = csvText.trim().split('\n');
  const headers = lines[0].split(',').map(h => h.trim());
  const videos = [];
  
  for (let i = 1; i < lines.length; i++) {
    const values = parseCSVLine(lines[i]);
    if (values.length === headers.length) {
      const video = {};
      headers.forEach((header, index) => {
        video[header] = values[index].trim();
      });
      videos.push(video);
    }
  }
  
  return videos;
}

function parseCSVLine(line) {
  const values = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      values.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  
  values.push(current);
  return values;
}

function processVideos() {
  const today = new Date();
  
  allVideos = allVideos.map(video => {
    // Generate ID if not exists
    if (!video.Video_ID) {
      video.Video_ID = generateID();
    }
    
    // Set default date if empty
    if (!video.Video_Date) {
      video.Video_Date = formatDate(today);
    }
    
    // Calculate attestation date (6 months from publish date)
    if (!video.Video_Attestation) {
      const publishDate = new Date(video.Video_Date);
      const attestationDate = new Date(publishDate);
      attestationDate.setMonth(attestationDate.getMonth() + 6);
      video.Video_Attestation = formatDate(attestationDate);
    }
    
    // Set default status
    if (!video.Video_Status) {
      video.Video_Status = 'Active';
    }
    
    // Calculate attestation status
    video.attestationStatus = calculateAttestationStatus(video.Video_Attestation);
    
    // Check if new (< 30 days)
    const publishDate = new Date(video.Video_Date);
    const daysSincePublish = Math.floor((today - publishDate) / (1000 * 60 * 60 * 24));
    video.isNew = daysSincePublish <= 30;
    
    return video;
  });
}

function calculateAttestationStatus(attestationDate) {
  const today = new Date();
  const attestDate = new Date(attestationDate);
  const daysUntil = Math.ceil((attestDate - today) / (1000 * 60 * 60 * 24));
  
  if (daysUntil < ATTESTATION_CRITICAL_DAYS) {
    return {
      status: 'critical',
      daysUntil: Math.abs(daysUntil),
      message: `Overdue by ${Math.abs(daysUntil)} days`
    };
  } else if (daysUntil <= ATTESTATION_WARNING_DAYS) {
    return {
      status: 'warning',
      daysUntil: daysUntil,
      message: `Review in ${daysUntil} days`
    };
  } else {
    return {
      status: 'success',
      daysUntil: daysUntil,
      message: `Up to date`
    };
  }
}

// ==========================================
// Filtering & Sorting
// ==========================================

function applyFiltersAndDisplay() {
  // Start with all active videos
  filteredVideos = allVideos.filter(v => v.Video_Status === 'Active');
  
  // Apply search filter
  if (searchQuery) {
    filteredVideos = filteredVideos.filter(video => {
      const searchLower = searchQuery.toLowerCase();
      return (
        video.Video_Title.toLowerCase().includes(searchLower) ||
        video.Video_Description.toLowerCase().includes(searchLower) ||
        video.Video_Author.toLowerCase().includes(searchLower)
      );
    });
  }
  
  // Apply category filter
  if (currentCategory) {
    filteredVideos = filteredVideos.filter(v => v.Video_Category === currentCategory);
  }
  
  // Apply sorting
  sortVideos();
  
  // Update UI
  updateAttestationWidget();
  updateResultsCount();
  updateClearFiltersButton();
  
  // Reset to page 1
  currentPage = 1;
  
  // Display videos
  displayVideos();
}

function sortVideos() {
  switch (currentSort) {
    case 'newest':
      filteredVideos.sort((a, b) => new Date(b.Video_Date) - new Date(a.Video_Date));
      break;
    case 'oldest':
      filteredVideos.sort((a, b) => new Date(a.Video_Date) - new Date(b.Video_Date));
      break;
    case 'title-asc':
      filteredVideos.sort((a, b) => a.Video_Title.localeCompare(b.Video_Title));
      break;
    case 'title-desc':
      filteredVideos.sort((a, b) => b.Video_Title.localeCompare(a.Video_Title));
      break;
    case 'attestation':
      filteredVideos.sort((a, b) => {
        return a.attestationStatus.daysUntil - b.attestationStatus.daysUntil;
      });
      break;
  }
}

function handleSearch(e) {
  const value = e.target.value;
  
  // Show/hide clear button
  document.getElementById('searchClear').style.display = value ? 'block' : 'none';
  
  // Debounce search
  clearTimeout(searchTimeout);
  searchTimeout = setTimeout(() => {
    if (value.length >= 2 || value.length === 0) {
      searchQuery = value;
      applyFiltersAndDisplay();
    }
  }, SEARCH_DEBOUNCE_DELAY);
}

function clearSearch() {
  document.getElementById('searchInput').value = '';
  document.getElementById('searchClear').style.display = 'none';
  searchQuery = '';
  applyFiltersAndDisplay();
}

function handleCategoryFilter(e) {
  currentCategory = e.target.value;
  applyFiltersAndDisplay();
}

function handleSort(e) {
  currentSort = e.target.value;
  applyFiltersAndDisplay();
}

function clearFilters() {
  // Reset filters
  searchQuery = '';
  currentCategory = '';
  currentSort = 'newest';
  
  // Reset UI
  document.getElementById('searchInput').value = '';
  document.getElementById('searchClear').style.display = 'none';
  document.getElementById('categoryFilter').value = '';
  document.getElementById('sortSelect').value = 'newest';
  
  // Reapply
  applyFiltersAndDisplay();
}

function updateClearFiltersButton() {
  const hasFilters = searchQuery || currentCategory || currentSort !== 'newest';
  document.getElementById('clearFiltersBtn').style.display = hasFilters ? 'block' : 'none';
}

function updateResultsCount() {
  const total = filteredVideos.length;
  const start = (currentPage - 1) * VIDEOS_PER_PAGE + 1;
  const end = Math.min(currentPage * VIDEOS_PER_PAGE, total);
  
  const countEl = document.getElementById('resultsCount');
  if (total === 0) {
    countEl.textContent = 'No videos found';
  } else if (total <= VIDEOS_PER_PAGE) {
    countEl.textContent = `Showing ${total} video${total !== 1 ? 's' : ''}`;
  } else {
    countEl.textContent = `Showing ${start}-${end} of ${total} videos`;
  }
  
  // Update badge count
  document.getElementById('videosCount').textContent = allVideos.filter(v => v.Video_Status === 'Active').length;
}

// ==========================================
// Display Videos
// ==========================================

function displayVideos() {
  const totalVideos = filteredVideos.length;
  
  if (totalVideos === 0) {
    showEmptyState('no-results');
    return;
  }
  
  hideEmptyState();
  
  // Calculate pagination
  const totalPages = Math.ceil(totalVideos / VIDEOS_PER_PAGE);
  const startIndex = (currentPage - 1) * VIDEOS_PER_PAGE;
  const endIndex = Math.min(startIndex + VIDEOS_PER_PAGE, totalVideos);
  const videosToDisplay = filteredVideos.slice(startIndex, endIndex);
  
  // Display based on view mode
  if (currentView === 'grid') {
    displayGridView(videosToDisplay);
  } else {
    displayListView(videosToDisplay);
  }
  
  // Update pagination
  updatePagination(totalPages);
  
  // Scroll to top
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function displayGridView(videos) {
  const gridEl = document.getElementById('videoGrid');
  const listEl = document.getElementById('videoList');
  
  gridEl.style.display = 'grid';
  listEl.style.display = 'none';
  
  gridEl.innerHTML = videos.map((video, index) => {
    return `
      <div class="vl-video-card" style="animation-delay: ${index * CARD_STAGGER_DELAY}ms">
        <div class="vl-video-thumbnail">
          ${video.Video_Embed_Code}
          ${video.Video_Duration ? `<div class="vl-badge-duration">${video.Video_Duration}</div>` : ''}
        </div>
        <div class="vl-video-content">
          <div class="vl-video-header">
            <h3 class="vl-video-title">${escapeHtml(video.Video_Title)}</h3>
          </div>
          
          <div class="vl-video-badges">
            ${video.isNew ? '<span class="vl-badge vl-badge-new">NEW</span>' : ''}
            <span class="vl-badge vl-badge-category">${escapeHtml(video.Video_Category)}</span>
            ${renderAttestationBadge(video.attestationStatus)}
          </div>
          
          <div class="vl-video-meta">
            <div class="vl-video-meta-item">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-user"></use></svg>
              <span>${escapeHtml(video.Video_Author)}</span>
            </div>
            <div class="vl-video-meta-item">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-calendar"></use></svg>
              <span>${formatDisplayDate(video.Video_Date)}</span>
            </div>
          </div>
          
          <p class="vl-video-description">${escapeHtml(video.Video_Description)}</p>
          
          <div class="vl-video-footer">
            <div class="vl-video-actions">
              ${video.Video_Document ? `
                <a href="${escapeHtml(video.Video_Document)}" target="_blank" rel="noopener noreferrer" class="vl-icon-btn" title="View related document">
                  <svg class="vl-icon vl-icon-sm"><use href="#icon-external"></use></svg>
                </a>
              ` : ''}
              ${editMode ? `
                <button class="vl-icon-btn" onclick="editVideo('${video.Video_ID}')" title="Edit video">
                  <svg class="vl-icon vl-icon-sm"><use href="#icon-edit"></use></svg>
                </button>
              ` : ''}
              ${deleteMode ? `
                <button class="vl-icon-btn danger" onclick="deleteVideo('${video.Video_ID}', '${escapeHtml(video.Video_Title)}')" title="Delete video">
                  <svg class="vl-icon vl-icon-sm"><use href="#icon-trash"></use></svg>
                </button>
              ` : ''}
            </div>
          </div>
        </div>
      </div>
    `;
  }).join('');
}

function displayListView(videos) {
  const gridEl = document.getElementById('videoGrid');
  const listEl = document.getElementById('videoList');
  
  gridEl.style.display = 'none';
  listEl.style.display = 'block';
  
  listEl.innerHTML = videos.map(video => {
    return `
      <div class="vl-list-item">
        <div class="vl-list-thumbnail">
          ${video.Video_Embed_Code}
        </div>
        
        <div class="vl-list-content">
          <h3 class="vl-list-title">${escapeHtml(video.Video_Title)}</h3>
          <p class="vl-list-description">${escapeHtml(video.Video_Description)}</p>
          
          <div class="vl-list-meta">
            <div class="vl-video-meta-item">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-user"></use></svg>
              <span>${escapeHtml(video.Video_Author)}</span>
            </div>
            <div class="vl-video-meta-item">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-calendar"></use></svg>
              <span>${formatDisplayDate(video.Video_Date)}</span>
            </div>
            <span class="vl-badge vl-badge-category">${escapeHtml(video.Video_Category)}</span>
            ${video.Video_Duration ? `<span class="vl-video-meta-item"><svg class="vl-icon vl-icon-sm"><use href="#icon-clock"></use></svg>${video.Video_Duration}</span>` : ''}
            ${renderAttestationBadge(video.attestationStatus)}
          </div>
        </div>
        
        <div class="vl-list-actions">
          ${video.Video_Document ? `
            <a href="${escapeHtml(video.Video_Document)}" target="_blank" rel="noopener noreferrer" class="vl-icon-btn" title="View related document">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-external"></use></svg>
            </a>
          ` : ''}
          ${editMode ? `
            <button class="vl-icon-btn" onclick="editVideo('${video.Video_ID}')" title="Edit video">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-edit"></use></svg>
            </button>
          ` : ''}
          ${deleteMode ? `
            <button class="vl-icon-btn danger" onclick="deleteVideo('${video.Video_ID}', '${escapeHtml(video.Video_Title)}')" title="Delete video">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-trash"></use></svg>
            </button>
          ` : ''}
        </div>
      </div>
    `;
  }).join('');
}

function renderAttestationBadge(attestationStatus) {
  if (attestationStatus.status === 'critical') {
    return `
      <span class="vl-badge vl-badge-danger">
        <svg class="vl-icon vl-icon-sm"><use href="#icon-error"></use></svg>
        ${attestationStatus.message}
      </span>
    `;
  } else if (attestationStatus.status === 'warning') {
    return `
      <span class="vl-badge vl-badge-warning">
        <svg class="vl-icon vl-icon-sm"><use href="#icon-warning"></use></svg>
        ${attestationStatus.message}
      </span>
    `;
  }
  return '';
}

// ==========================================
// Pagination
// ==========================================

function updatePagination(totalPages) {
  const paginationEl = document.getElementById('pagination');
  const pagesEl = document.getElementById('paginationPages');
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  
  if (totalPages <= 1) {
    paginationEl.style.display = 'none';
    return;
  }
  
  paginationEl.style.display = 'flex';
  
  // Update prev/next buttons
  prevBtn.disabled = currentPage === 1;
  nextBtn.disabled = currentPage === totalPages;
  
  // Generate page numbers
  const pages = generatePageNumbers(currentPage, totalPages);
  pagesEl.innerHTML = pages.map(page => {
    if (page === '...') {
      return '<span class="vl-page-ellipsis">...</span>';
    }
    return `
      <button class="vl-page-btn ${page === currentPage ? 'active' : ''}" 
              onclick="changePage(${page})">
        ${page}
      </button>
    `;
  }).join('');
}

function generatePageNumbers(current, total) {
  const pages = [];
  
  if (total <= 7) {
    for (let i = 1; i <= total; i++) {
      pages.push(i);
    }
  } else {
    if (current <= 3) {
      pages.push(1, 2, 3, 4, '...', total);
    } else if (current >= total - 2) {
      pages.push(1, '...', total - 3, total - 2, total - 1, total);
    } else {
      pages.push(1, '...', current - 1, current, current + 1, '...', total);
    }
  }
  
  return pages;
}

function changePage(page) {
  currentPage = page;
  displayVideos();
}

// ==========================================
// View Toggle
// ==========================================

function setView(view) {
  currentView = view;
  
  // Update buttons
  document.getElementById('gridViewBtn').classList.toggle('active', view === 'grid');
  document.getElementById('listViewBtn').classList.toggle('active', view === 'list');
  
  // Save preference
  localStorage.setItem(STORAGE_VIEW_MODE, view);
  
  // Redisplay
  displayVideos();
}

// ==========================================
// Attestation Widget
// ==========================================

function updateAttestationWidget() {
  const upToDate = allVideos.filter(v => 
    v.Video_Status === 'Active' && v.attestationStatus.status === 'success'
  ).length;
  
  const warning = allVideos.filter(v => 
    v.Video_Status === 'Active' && v.attestationStatus.status === 'warning'
  ).length;
  
  const overdue = allVideos.filter(v => 
    v.Video_Status === 'Active' && v.attestationStatus.status === 'critical'
  ).length;
  
  document.getElementById('statUpToDate').textContent = upToDate;
  document.getElementById('statWarning').textContent = warning;
  document.getElementById('statOverdue').textContent = overdue;
}

function toggleAttestationWidget() {
  const content = document.getElementById('widgetContent');
  const toggle = document.getElementById('widgetToggle');
  
  content.classList.toggle('collapsed');
  toggle.classList.toggle('collapsed');
}

function toggleAttestationDashboardVisibility() {
  const widget = document.getElementById('attestationWidget');
  const toggleBtn = document.getElementById('toggleAttestationDashboard');
  const isVisible = widget.style.display !== 'none';
  
  if (isVisible) {
    widget.style.display = 'none';
    toggleBtn.querySelector('span').textContent = 'Show Attestation Dashboard';
    localStorage.setItem('videoLibrary_showAttestation', 'false');
  } else {
    widget.style.display = 'block';
    toggleBtn.querySelector('span').textContent = 'Hide Attestation Dashboard';
    localStorage.setItem('videoLibrary_showAttestation', 'true');
  }
  
  // Close dropdown
  document.getElementById('configMenu').classList.remove('show');
}

// ==========================================
// Modes (Edit/Delete)
// ==========================================

function toggleEditMode() {
  editMode = !editMode;
  deleteMode = false; // Only one mode at a time
  displayVideos();
  showToast('info', editMode ? 'Edit mode enabled' : 'Edit mode disabled');
}

function toggleDeleteMode() {
  deleteMode = !deleteMode;
  editMode = false; // Only one mode at a time
  displayVideos();
  showToast('info', deleteMode ? 'Delete mode enabled' : 'Delete mode disabled');
}

// ==========================================
// Video CRUD Operations
// ==========================================

function openVideoModal(videoId = null) {
  const modal = document.getElementById('videoModal');
  const title = document.getElementById('modalTitle');
  const form = document.getElementById('videoForm');
  
  // Reset form
  form.reset();
  clearFormErrors();
  document.getElementById('titleCounter').textContent = '0';
  document.getElementById('descCounter').textContent = '0';
  
  if (videoId) {
    // Edit mode
    const video = allVideos.find(v => v.Video_ID === videoId);
    if (video) {
      title.textContent = `Edit Video: ${video.Video_Title}`;
      populateForm(video);
    }
  } else {
    // Add mode
    title.textContent = 'Add New Video';
    document.getElementById('videoDate').value = formatDate(new Date());
  }
  
  modal.style.display = 'flex';
  document.getElementById('videoTitle').focus();
}

function populateForm(video) {
  document.getElementById('videoId').value = video.Video_ID;
  document.getElementById('videoTitle').value = video.Video_Title;
  document.getElementById('videoDescription').value = video.Video_Description;
  document.getElementById('videoAuthor').value = video.Video_Author;
  document.getElementById('videoDate').value = video.Video_Date;
  document.getElementById('videoCategory').value = video.Video_Category;
  document.getElementById('videoDocument').value = video.Video_Document || '';
  document.getElementById('videoEmbed').value = video.Video_Embed_Code;
  document.getElementById('videoDuration').value = video.Video_Duration || '';
  document.getElementById('videoThumbnail').value = video.Video_Thumbnail || '';
  
  // Update counters
  document.getElementById('titleCounter').textContent = video.Video_Title.length;
  document.getElementById('descCounter').textContent = video.Video_Description.length;
}

function closeVideoModal() {
  // Check for unsaved changes
  const form = document.getElementById('videoForm');
  const hasChanges = Array.from(form.elements).some(el => 
    el.type !== 'submit' && el.type !== 'button' && el.value
  );
  
  if (hasChanges) {
    if (!confirm('You have unsaved changes. Are you sure you want to close?')) {
      return;
    }
  }
  
  document.getElementById('videoModal').style.display = 'none';
}

async function handleVideoSubmit(e) {
  e.preventDefault();
  
  // Validate form
  if (!validateForm()) {
    return;
  }
  
  const saveBtn = document.getElementById('modalSave');
  const saveText = document.getElementById('saveText');
  const saveSpinner = document.getElementById('saveSpinner');
  
  // Show loading
  saveBtn.disabled = true;
  saveText.style.display = 'none';
  saveSpinner.style.display = 'inline-block';
  
  try {
    const videoData = getFormData();
    const isEdit = !!videoData.Video_ID;
    
    if (isEdit) {
      await updateVideo(videoData);
    } else {
      await createVideo(videoData);
    }
    
    // Reload videos
    await loadVideos();
    
    // Close modal
    document.getElementById('videoModal').style.display = 'none';
    
    // Show success
    showToast('success', isEdit ? 'Video updated successfully' : 'Video added successfully');
  } catch (error) {
    console.error('Error saving video:', error);
    showToast('error', 'Failed to save video. Please try again.');
  } finally {
    // Hide loading
    saveBtn.disabled = false;
    saveText.style.display = 'inline';
    saveSpinner.style.display = 'none';
  }
}

function getFormData() {
  const videoId = document.getElementById('videoId').value;
  
  return {
    Video_ID: videoId || generateID(),
    Video_Title: document.getElementById('videoTitle').value.trim(),
    Video_Description: document.getElementById('videoDescription').value.trim(),
    Video_Author: document.getElementById('videoAuthor').value.trim(),
    Video_Date: document.getElementById('videoDate').value || formatDate(new Date()),
    Video_Category: document.getElementById('videoCategory').value,
    Video_Document: document.getElementById('videoDocument').value.trim(),
    Video_Embed_Code: document.getElementById('videoEmbed').value.trim(),
    Video_Duration: document.getElementById('videoDuration').value.trim(),
    Video_Thumbnail: document.getElementById('videoThumbnail').value.trim(),
    Video_Status: 'Active'
  };
}

async function createVideo(videoData) {
  // Calculate attestation date
  const publishDate = new Date(videoData.Video_Date);
  const attestationDate = new Date(publishDate);
  attestationDate.setMonth(attestationDate.getMonth() + 6);
  videoData.Video_Attestation = formatDate(attestationDate);
  
  // Add to array
  allVideos.push(videoData);
  
  // Save to CSV
  await saveVideosToCSV();
  
  // Log audit
  await logAudit('Created', videoData.Video_ID, videoData.Video_Title, null);
}

async function updateVideo(videoData) {
  const index = allVideos.findIndex(v => v.Video_ID === videoData.Video_ID);
  if (index === -1) return;
  
  const oldVideo = { ...allVideos[index] };
  
  // Recalculate attestation if date changed
  if (oldVideo.Video_Date !== videoData.Video_Date) {
    const publishDate = new Date(videoData.Video_Date);
    const attestationDate = new Date(publishDate);
    attestationDate.setMonth(attestationDate.getMonth() + 6);
    videoData.Video_Attestation = formatDate(attestationDate);
  } else {
    videoData.Video_Attestation = oldVideo.Video_Attestation;
  }
  
  // Update array
  allVideos[index] = videoData;
  
  // Save to CSV
  await saveVideosToCSV();
  
  // Log audit with changes
  const changes = getChanges(oldVideo, videoData);
  await logAudit('Updated', videoData.Video_ID, videoData.Video_Title, changes);
}

let videoToDelete = null;

function deleteVideo(videoId, videoTitle) {
  videoToDelete = { id: videoId, title: videoTitle };
  document.getElementById('deleteVideoTitle').textContent = videoTitle;
  document.getElementById('deleteModal').style.display = 'flex';
}

function closeDeleteModal() {
  videoToDelete = null;
  document.getElementById('deleteModal').style.display = 'none';
}

async function confirmDelete() {
  if (!videoToDelete) return;
  
  const deleteBtn = document.getElementById('deleteConfirm');
  const deleteText = document.getElementById('deleteText');
  const deleteSpinner = document.getElementById('deleteSpinner');
  
  // Show loading
  deleteBtn.disabled = true;
  deleteText.style.display = 'none';
  deleteSpinner.style.display = 'inline-block';
  
  try {
    const { id, title } = videoToDelete;
    
    // Remove from array
    allVideos = allVideos.filter(v => v.Video_ID !== id);
    
    // Save to CSV
    await saveVideosToCSV();
    
    // Log audit
    await logAudit('Deleted', id, title, null);
    
    // Reload videos
    await loadVideos();
    
    // Close modal
    document.getElementById('deleteModal').style.display = 'none';
    videoToDelete = null;
    
    // Show success
    showToast('success', 'Video deleted successfully');
  } catch (error) {
    console.error('Error deleting video:', error);
    showToast('error', 'Failed to delete video. Please try again.');
  } finally {
    // Hide loading
    deleteBtn.disabled = false;
    deleteText.style.display = 'inline';
    deleteSpinner.style.display = 'none';
  }
}

function editVideo(videoId) {
  openVideoModal(videoId);
}

// ==========================================
// CSV Operations
// ==========================================

async function saveVideosToCSV() {
  const headers = [
    'Video_ID', 'Video_Title', 'Video_Description', 'Video_Author',
    'Video_Date', 'Video_Category', 'Video_Document', 'Video_Attestation',
    'Video_Embed_Code', 'Video_Duration', 'Video_Thumbnail', 'Video_Status'
  ];
  
  const csvLines = [headers.join(',')];
  
  allVideos.forEach(video => {
    const row = headers.map(header => {
      const value = video[header] || '';
      // Escape quotes and wrap in quotes if contains comma
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        return `"${value.replace(/"/g, '""')}"`;
      }
      return value;
    });
    csvLines.push(row.join(','));
  });
  
  const csvContent = csvLines.join('\n');
  
  // Upload to Confluence
  await uploadCSVToConfluence(CSV_FILENAME_VIDEOS, csvContent);
}

async function uploadCSVToConfluence(filename, content) {
  try {
    const formData = new FormData();
    const blob = new Blob([content], { type: 'text/csv' });
    formData.append('file', blob, filename);
    formData.append('comment', `Updated via Video Library App`);
    
    const url = `${CONFLUENCE_BASE_URL}/rest/api/content/${CONFLUENCE_PAGE_ID}/child/attachment`;
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'X-Atlassian-Token': 'no-check'
      },
      body: formData
    });
    
    if (!response.ok) {
      throw new Error('Failed to upload file');
    }
    
    return await response.json();
  } catch (error) {
    console.error('Error uploading CSV:', error);
    throw error;
  }
}

// ==========================================
// Audit Log
// ==========================================

async function logAudit(action, videoId, videoTitle, changes) {
  try {
    // Fetch current audit log
    let auditLog = [];
    try {
      const csvData = await fetchCSVFromConfluence(CSV_FILENAME_AUDIT);
      if (csvData) {
        auditLog = parseCSV(csvData);
      }
    } catch (error) {
      console.log('No existing audit log, creating new one');
    }
    
    // Create new log entry
    const timestamp = new Date().toISOString();
    const logEntry = {
      Log_ID: timestamp.replace(/[:.]/g, '-'),
      Timestamp: timestamp,
      User_Name: await getConfluenceUser(),
      Action: action,
      Video_ID: videoId,
      Video_Title: videoTitle,
      Changes_Made: changes ? JSON.stringify(changes) : '',
      IP_Address: ''
    };
    
    auditLog.push(logEntry);
    
    // Convert to CSV
    const headers = [
      'Log_ID', 'Timestamp', 'User_Name', 'Action',
      'Video_ID', 'Video_Title', 'Changes_Made', 'IP_Address'
    ];
    
    const csvLines = [headers.join(',')];
    auditLog.forEach(log => {
      const row = headers.map(header => {
        const value = log[header] || '';
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          return `"${value.replace(/"/g, '""')}"`;
        }
        return value;
      });
      csvLines.push(row.join(','));
    });
    
    const csvContent = csvLines.join('\n');
    
    // Upload
    await uploadCSVToConfluence(CSV_FILENAME_AUDIT, csvContent);
  } catch (error) {
    console.error('Error logging audit:', error);
    // Don't throw - audit log failure shouldn't break main functionality
  }
}

async function getConfluenceUser() {
  try {
    const response = await fetch(`${CONFLUENCE_BASE_URL}/rest/api/user/current`);
    if (response.ok) {
      const user = await response.json();
      return user.displayName || user.username || 'Unknown User';
    }
  } catch (error) {
    console.error('Error fetching user:', error);
  }
  return 'Unknown User';
}

function getChanges(oldVideo, newVideo) {
  const changes = {};
  const fields = [
    'Video_Title', 'Video_Description', 'Video_Author', 'Video_Date',
    'Video_Category', 'Video_Document', 'Video_Embed_Code',
    'Video_Duration', 'Video_Thumbnail'
  ];
  
  fields.forEach(field => {
    if (oldVideo[field] !== newVideo[field]) {
      changes[field] = {
        old: oldVideo[field],
        new: newVideo[field]
      };
    }
  });
  
  return changes;
}

// ==========================================
// Export Functions
// ==========================================

async function exportVideos() {
  try {
    const csvContent = await generateVideosCSV(filteredVideos);
    downloadCSV(csvContent, `training_videos_export_${formatDate(new Date())}.csv`);
    showToast('success', 'Videos exported successfully');
    
    // Log audit
    await logAudit('Exported', 'N/A', `${filteredVideos.length} videos`, null);
  } catch (error) {
    console.error('Error exporting videos:', error);
    showToast('error', 'Failed to export videos');
  }
}

async function exportAuditLog() {
  try {
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_AUDIT);
    if (csvData) {
      downloadCSV(csvData, `video_audit_log_${formatDate(new Date())}.csv`);
      showToast('success', 'Audit log exported successfully');
      
      // Log the export
      await logAudit('Exported', 'N/A', 'Audit Log', null);
    } else {
      showToast('warning', 'No audit log found');
    }
  } catch (error) {
    console.error('Error exporting audit log:', error);
    showToast('error', 'Failed to export audit log');
  }
}

function generateVideosCSV(videos) {
  const headers = [
    'Video_ID', 'Video_Title', 'Video_Description', 'Video_Author',
    'Video_Date', 'Video_Category', 'Video_Document', 'Video_Attestation',
    'Video_Duration', 'Video_Thumbnail', 'Video_Status'
  ];
  
  const csvLines = [headers.join(',')];
  
  videos.forEach(video => {
    const row = headers.map(header => {
      const value = video[header] || '';
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        return `"${value.replace(/"/g, '""')}"`;
      }
      return value;
    });
    csvLines.push(row.join(','));
  });
  
  return csvLines.join('\n');
}

function downloadCSV(content, filename) {
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';
  
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// ==========================================
// Form Validation
// ==========================================

function validateForm() {
  let isValid = true;
  
  const titleInput = document.getElementById('videoTitle');
  const descInput = document.getElementById('videoDescription');
  const authorInput = document.getElementById('videoAuthor');
  const categoryInput = document.getElementById('videoCategory');
  const embedInput = document.getElementById('videoEmbed');
  
  if (!validateField(titleInput, 'titleError', 'Title is required')) isValid = false;
  if (!validateField(descInput, 'descError', 'Description is required')) isValid = false;
  if (!validateField(authorInput, 'authorError', 'Author name is required')) isValid = false;
  if (!validateField(categoryInput, 'categoryError', 'Please select a category')) isValid = false;
  if (!validateEmbedCode(embedInput)) isValid = false;
  
  const durationInput = document.getElementById('videoDuration');
  if (durationInput.value && !validateDuration(durationInput)) isValid = false;
  
  const documentInput = document.getElementById('videoDocument');
  if (documentInput.value && !validateURL(documentInput, 'documentError')) isValid = false;
  
  const thumbnailInput = document.getElementById('videoThumbnail');
  if (thumbnailInput.value && !validateURL(thumbnailInput, 'thumbnailError')) isValid = false;
  
  return isValid;
}

function validateField(input, errorId, message) {
  const errorEl = document.getElementById(errorId);
  
  if (!input.value.trim()) {
    input.classList.add('error');
    errorEl.textContent = message;
    return false;
  }
  
  input.classList.remove('error');
  errorEl.textContent = '';
  return true;
}

function validateEmbedCode(input) {
  const errorEl = document.getElementById('embedError');
  const value = input.value.trim();
  
  if (!value) {
    input.classList.add('error');
    errorEl.textContent = 'Embed code is required';
    return false;
  }
  
  if (!value.includes('<iframe')) {
    input.classList.add('error');
    errorEl.textContent = 'Please provide a valid iframe embed code';
    return false;
  }
  
  input.classList.remove('error');
  errorEl.textContent = '';
  return true;
}

function validateDuration(input) {
  const errorEl = document.getElementById('durationError');
  const value = input.value.trim();
  
  if (!value) {
    input.classList.remove('error');
    errorEl.textContent = '';
    return true;
  }
  
  const pattern = /^[0-5]?[0-9]:[0-5][0-9]$/;
  if (!pattern.test(value)) {
    input.classList.add('error');
    errorEl.textContent = 'Format should be MM:SS (e.g., 05:30)';
    return false;
  }
  
  input.classList.remove('error');
  errorEl.textContent = '';
  return true;
}

function validateURL(input, errorId) {
  const errorEl = document.getElementById(errorId);
  const value = input.value.trim();
  
  if (!value) {
    input.classList.remove('error');
    errorEl.textContent = '';
    return true;
  }
  
  try {
    new URL(value);
    input.classList.remove('error');
    errorEl.textContent = '';
    return true;
  } catch (e) {
    input.classList.add('error');
    errorEl.textContent = 'Please enter a valid URL';
    return false;
  }
}

function clearFormErrors() {
  const errorElements = document.querySelectorAll('.vl-error');
  errorElements.forEach(el => el.textContent = '');
  
  const inputs = document.querySelectorAll('.vl-input, .vl-textarea');
  inputs.forEach(input => input.classList.remove('error'));
}

// ==========================================
// UI Helpers
// ==========================================

function showSkeleton() {
  document.getElementById('skeletonContainer').style.display = 'grid';
  document.getElementById('videoGrid').style.display = 'none';
  document.getElementById('videoList').style.display = 'none';
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('pagination').style.display = 'none';
}

function hideSkeleton() {
  document.getElementById('skeletonContainer').style.display = 'none';
}

function showEmptyState(type) {
  const emptyState = document.getElementById('emptyState');
  const title = document.getElementById('emptyTitle');
  const text = document.getElementById('emptyText');
  const actionBtn = document.getElementById('emptyActionBtn');
  
  document.getElementById('videoGrid').style.display = 'none';
  document.getElementById('videoList').style.display = 'none';
  document.getElementById('pagination').style.display = 'none';
  emptyState.style.display = 'block';
  
  if (type === 'no-results') {
    title.textContent = 'No videos found';
    text.textContent = 'Try adjusting your filters or search terms';
    actionBtn.style.display = 'none';
  } else if (type === 'error') {
    title.textContent = 'Unable to load videos';
    text.textContent = 'Please refresh the page to try again';
    actionBtn.style.display = 'none';
  } else {
    title.textContent = 'No training videos yet';
    text.textContent = 'Get started by adding your first video';
    actionBtn.style.display = 'inline-flex';
  }
}

function hideEmptyState() {
  document.getElementById('emptyState').style.display = 'none';
}

function toggleDropdown(menu) {
  menu.classList.toggle('show');
}

function showToast(type, message) {
  const container = document.getElementById('toastContainer');
  
  const iconMap = {
    success: 'check',
    error: 'error',
    warning: 'warning',
    info: 'info'
  };
  
  const toast = document.createElement('div');
  toast.className = `vl-toast vl-toast-${type}`;
  toast.innerHTML = `
    <div class="vl-toast-icon">
      <svg class="vl-icon vl-icon-md"><use href="#icon-${iconMap[type]}"></use></svg>
    </div>
    <div class="vl-toast-content">
      <p class="vl-toast-message">${escapeHtml(message)}</p>
    </div>
    <button class="vl-toast-close">
      <svg class="vl-icon vl-icon-sm"><use href="#icon-close"></use></svg>
    </button>
  `;
  
  // Close button
  toast.querySelector('.vl-toast-close').addEventListener('click', () => {
    removeToast(toast);
  });
  
  container.appendChild(toast);
  
  // Auto remove
  setTimeout(() => {
    removeToast(toast);
  }, TOAST_DURATION);
}

function removeToast(toast) {
  toast.classList.add('removing');
  setTimeout(() => {
    if (toast.parentElement) {
      toast.parentElement.removeChild(toast);
    }
  }, 300);
}

function populateCategoryDropdown() {
  const categoryFilter = document.getElementById('categoryFilter');
  const categorySelect = document.getElementById('videoCategory');
  
  VIDEO_CATEGORIES.forEach(category => {
    const option = document.createElement('option');
    option.value = category;
    option.textContent = category;
    
    categoryFilter.appendChild(option.cloneNode(true));
    categorySelect.appendChild(option);
  });
}

// ==========================================
// Utility Functions
// ==========================================

function generateID() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function formatDate(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDisplayDate(dateStr) {
  const date = new Date(dateStr);
  const options = { year: 'numeric', month: 'short', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, m => map[m]);
}

function loadPreferences() {
  const savedView = localStorage.getItem(STORAGE_VIEW_MODE);
  if (savedView) {
    currentView = savedView;
    document.getElementById('gridViewBtn').classList.toggle('active', savedView === 'grid');
    document.getElementById('listViewBtn').classList.toggle('active', savedView === 'list');
  }
  
  // Load attestation dashboard visibility preference
  const showAttestation = localStorage.getItem('videoLibrary_showAttestation');
  const widget = document.getElementById('attestationWidget');
  const toggleBtn = document.getElementById('toggleAttestationDashboard');
  
  if (showAttestation === 'true') {
    widget.style.display = 'block';
    if (toggleBtn) {
      toggleBtn.querySelector('span').textContent = 'Hide Attestation Dashboard';
    }
  } else {
    widget.style.display = 'none';
    if (toggleBtn) {
      toggleBtn.querySelector('span').textContent = 'Show Attestation Dashboard';
    }
  }
}
