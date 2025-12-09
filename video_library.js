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

const CONFLUENCE_BASE_URL = 'mysite/wiki';
const CONFLUENCE_PAGE_ID = '1235678990';
const CSV_FILENAME_VIDEOS = 'training_videos.csv';
const CSV_FILENAME_AUDIT = 'video_audit_log.csv';
const CSV_FILENAME_STATS = 'video_stats.csv';

// Default thumbnail for videos without custom thumbnail
const DEFAULT_THUMBNAIL_URL = 'https://via.placeholder.com/640x360/f0f2f5/305edb?text=Video+Thumbnail';

// Dynamic category colors - will be assigned based on unique categories from CSV
const CATEGORY_COLORS = [
  '#305edb', // Citi Blue (primary)
  '#2a6b3c', // Green
  '#fab728', // Orange
  '#9a231a', // Red
  '#6366f1', // Indigo
  '#8b5cf6', // Purple
  '#ec4899', // Pink
  '#14b8a6', // Teal
  '#f59e0b', // Amber
  '#06b6d4'  // Cyan
];

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
let userAvatarCache = {}; // Cache for user avatars
let dynamicCategories = []; // Dynamically extracted from CSV
let categoryColorMap = {}; // Maps category names to colors
let currentUserName = null;
let videoStatsRecords = [];
let videoStats = { likes: {}, watches: {} };
let focusVideoId = null;

// Initialize Application
document.addEventListener('DOMContentLoaded', async () => {
  await initializeApp();
  
  // Check if URL has a video ID parameter for direct linking
  const urlParams = new URLSearchParams(window.location.search);
  const videoId = urlParams.get('video');
  if (videoId) {
    // Wait a moment for videos to load
    setTimeout(() => {
      const video = allVideos.find(v => v.Video_ID === videoId);
      if (video) {
        handleWatch(videoId);
      }
    }, 1000);
  }
});

async function initializeApp() {
  // Load saved preferences
  loadPreferences();
  
  // Setup event listeners
  setupEventListeners();
  
  // Populate category dropdown
  populateCategoryDropdown();
  
  // Load videos
  await loadCurrentUser();
  await loadVideoStats();
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
  
  // Refresh button
  document.getElementById('refreshBtn').addEventListener('click', handleRefresh);
  
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
  document.getElementById('attestationManagerBtn').addEventListener('click', () => {
    toggleDropdown(configMenu);
    openAttestationManager();
  });
  const focusClose = document.getElementById('focusCloseBtn');
  if (focusClose) {
    focusClose.addEventListener('click', closeFocusView);
  }
  
  // Filters
  document.getElementById('categoryFilter').addEventListener('change', handleCategoryFilter);
  document.getElementById('sortSelect').addEventListener('change', handleSort);
  document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
  
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
  
  // Attestation modal
  document.getElementById('attestationModalClose').addEventListener('click', closeAttestationModal);
  document.getElementById('attestationCancel').addEventListener('click', closeAttestationModal);
  document.getElementById('attestationSave').addEventListener('click', saveAttestationExtension);
  
  // Empty state action
  document.getElementById('emptyActionBtn').addEventListener('click', () => openVideoModal());
  
  // Close modals on overlay click
  document.getElementById('videoModal').addEventListener('click', (e) => {
    if (e.target.id === 'videoModal') closeVideoModal();
  });
  document.getElementById('deleteModal').addEventListener('click', (e) => {
    if (e.target.id === 'deleteModal') closeDeleteModal();
  });
  document.getElementById('attestationModal').addEventListener('click', (e) => {
    if (e.target.id === 'attestationModal') closeAttestationModal();
  });
  
  // ESC key to close modals
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeVideoModal();
      closeDeleteModal();
      closeAttestationModal();
    }
  });
  
  updateModeButtons();
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
  
  document.getElementById('videoAuthorUsername').addEventListener('blur', (e) => {
    validateField(e.target, 'authorUsernameError', 'Author username is required');
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
    console.log('=== LOADING VIDEOS ===');
    showSkeleton();
    
    // Fetch CSV from Confluence
    console.log('Fetching CSV from Confluence...');
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_VIDEOS);
    console.log('CSV data received. Length:', csvData.length);
    
    if (!csvData) {
      console.log('No CSV data found - first time setup');
      allVideos = [];
    } else {
      // Parse CSV to JSON
      console.log('Parsing CSV...');
      allVideos = parseCSV(csvData);
      console.log('Parsed videos count:', allVideos.length);
      
      if (allVideos.length > 0) {
        console.log('Sample video:', {
          id: allVideos[0].Video_ID,
          title: allVideos[0].Video_Title,
          author: allVideos[0].Video_Author
        });
      }
    }
    
    // Process videos (calculate attestation, etc.)
    console.log('Processing videos...');
    processVideos();
    
    // Apply filters and display
    console.log('Applying filters and displaying...');
    applyFiltersAndDisplay();
    
    hideSkeleton();
    console.log('=== LOADING COMPLETE ===');
  } catch (error) {
    console.error('=== LOADING ERROR ===');
    console.error('Error loading videos:', error);
    console.error('Error stack:', error.stack);
    hideSkeleton();
    showToast('error', `Failed to load videos: ${error.message}`);
    showEmptyState('error');
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
      console.log('File not found - first time use');
      return '';
    }
    
    console.log('File found:', {
      id: attachment.id,
      title: attachment.title,
      version: attachment.version?.number
    });
    
    // Download the file
    const downloadUrl = `${CONFLUENCE_BASE_URL}${attachment._links.download}`;
    console.log('Download URL:', downloadUrl);
    
    const fileResponse = await fetch(downloadUrl, {
      credentials: 'include'
    });
    
    console.log('Download response status:', fileResponse.status);
    
    if (!fileResponse.ok) {
      const errorText = await fileResponse.text();
      console.error('Failed to download file:', errorText);
      throw new Error(`Failed to download file: ${fileResponse.status} ${fileResponse.statusText}`);
    }
    
    const csvContent = await fileResponse.text();
    console.log('CSV downloaded. Length:', csvContent.length);
    console.log('First 300 chars:', csvContent.substring(0, 300));
    
    return csvContent;
  } catch (error) {
    console.error('=== FETCH ERROR ===');
    console.error('Error fetching CSV:', error);
    throw error;
  }
}

function parseCSV(csvText) {
  console.log('=== PARSING CSV ===');
  
  if (!csvText || csvText.trim() === '') {
    console.log('Empty CSV text');
    return [];
  }
  
  const lines = csvText.trim().split('\n');
  console.log('Total lines:', lines.length);
  
  if (lines.length === 0) {
    console.log('No lines in CSV');
    return [];
  }
  
  const headers = parseCSVLine(lines[0]).map(h => h.trim());
  console.log('Headers:', headers);
  
  const videos = [];
  
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) {
      console.log(`Line ${i + 1}: Empty, skipping`);
      continue;
    }
    
    const values = parseCSVLine(lines[i]);
    
    if (values.length > 0) {
      const video = {};
      headers.forEach((header, index) => {
        video[header] = values[index] ? values[index].trim() : '';
      });
      videos.push(video);
      
      if (i <= 3) {
        console.log(`Line ${i + 1} parsed:`, {
          id: video.Video_ID,
          title: video.Video_Title
        });
      }
    }
  }
  
  console.log('Total videos parsed:', videos.length);
  return videos;
}

function parseCSVLine(line) {
  const values = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = line[i + 1];
    
    if (char === '"' && nextChar === '"' && inQuotes) {
      // Handle escaped quotes
      current += '"';
      i++; // Skip next quote
    } else if (char === '"') {
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
  
  // Extract unique categories dynamically from CSV
  extractDynamicCategories();
  
  // Populate category dropdown with dynamic categories
  populateCategoryDropdown();
  
  // Pre-fetch all avatars to prevent loading issues
  allVideos.forEach(video => {
    if (video.Video_Author_Username) {
      fetchUserAvatar(video.Video_Author_Username);
    }
  });
}

function extractDynamicCategories() {
  // Extract unique categories from all videos
  const categorySet = new Set();
  allVideos.forEach(video => {
    if (video.Video_Category && video.Video_Category.trim()) {
      categorySet.add(video.Video_Category.trim());
    }
  });
  
  // Convert to sorted array
  dynamicCategories = Array.from(categorySet).sort();
  
  // Assign colors to each category
  dynamicCategories.forEach((category, index) => {
    categoryColorMap[category] = CATEGORY_COLORS[index % CATEGORY_COLORS.length];
  });
  
  console.log('Dynamic categories extracted:', dynamicCategories);
  console.log('Category color map:', categoryColorMap);
}

function getCategoryColor(category) {
  return categoryColorMap[category] || CATEGORY_COLORS[0];
}

function getCategoryStyle(category) {
  const color = getCategoryColor(category);
  
  return {
    backgroundColor: color,
    color: '#ffffff',
    borderColor: color
  };
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
  updateResultsCount();
  updateClearFiltersButton();
  
  // Reset to page 1
  currentPage = 1;
  
  // If in focus mode, only update the rail, don't show grid/list
  if (focusVideoId) {
    buildFocusRail(focusVideoId);
  } else {
    // Display videos normally
    displayVideos();
  }
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

async function handleRefresh() {
  const btn = document.getElementById('refreshBtn');
  btn.classList.add('loading');
  
  try {
    // If in focus mode, close it before refreshing
    if (focusVideoId) {
      closeFocusView();
    }
    
    await loadVideoStats();
    await loadVideos();
    showToast('success', 'Data refreshed successfully');
  } catch (error) {
    console.error('Error refreshing data:', error);
    showToast('error', 'Failed to refresh data');
  } finally {
    btn.classList.remove('loading');
  }
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

  // Refresh focus rail if open
  const focusContainer = document.getElementById('focusView');
  if (focusContainer && focusContainer.style.display !== 'none' && focusVideoId) {
    buildFocusRail(focusVideoId);
  }
}

function displayGridView(videos) {
  const gridEl = document.getElementById('videoGrid');
  const listEl = document.getElementById('videoList');
  
  gridEl.style.display = 'grid';
  listEl.style.display = 'none';
  
  // Fetch avatars for all videos before rendering
  const avatarPromises = videos.map(video => {
    if (video.Video_Author_Username) {
      return fetchUserAvatar(video.Video_Author_Username);
    }
    return Promise.resolve();
  });
  
  // Wait a bit for avatars to load, then render
  Promise.all(avatarPromises).finally(() => {
    gridEl.innerHTML = videos.map((video, index) => renderAltCard(video, index)).join('');
  });
}

function renderAltCard(video, index) {
  const avatarUrl = userAvatarCache[video.Video_Author_Username] || '';
  const categoryStyle = getCategoryStyle(video.Video_Category);
  const stats = getVideoStats(video.Video_ID);
  const thumb = getThumbnailUrl(video);
  
  return `
    <div class="vl-video-card alt-card" style="animation-delay: ${index * CARD_STAGGER_DELAY}ms">
      <div class="vl-video-thumbnail">
        ${thumb ? `<img src="${thumb}" alt="${escapeHtml(video.Video_Title)}"/>` : ''}
        <div class="vl-card-overlay" onclick="handleWatch('${video.Video_ID}')"></div>
        ${video.Video_Duration ? `<div class="vl-badge-duration">${video.Video_Duration}</div>` : ''}
      </div>
      <div class="vl-video-content">
        <div class="vl-video-header">
          <h3 class="vl-video-title">${escapeHtml(video.Video_Title)}</h3>
        </div>
        
        <div class="vl-video-meta">
          <div class="vl-video-meta-item">
            ${avatarUrl ? `<img src="${avatarUrl}" alt="${escapeHtml(video.Video_Author)}" class="vl-user-avatar" />` : '<svg class="vl-icon vl-icon-sm"><use href="#icon-user"></use></svg>'}
            <span>${escapeHtml(video.Video_Author)}</span>
          </div>
          <div class="vl-video-meta-item">
            <svg class="vl-icon vl-icon-sm"><use href="#icon-calendar"></use></svg>
            <span>${formatDisplayDate(video.Video_Date)}</span>
          </div>
          ${video.Video_Duration ? `<span class="vl-video-meta-item"><svg class="vl-icon vl-icon-sm"><use href="#icon-clock"></use></svg>${video.Video_Duration}</span>` : ''}
        </div>
        
        <div class="vl-video-badges">
          <div style="display: flex; gap: 8px; align-items: center;">
            <span class="vl-badge vl-badge-category" style="background: ${categoryStyle.backgroundColor}; color: ${categoryStyle.color}; border: 1px solid ${categoryStyle.borderColor};">
              ${escapeHtml(video.Video_Category)}
            </span>
            ${renderAttestationBadge(video)}
          </div>
          ${video.isNew ? '<span class="vl-badge vl-badge-new"><svg class="vl-icon vl-icon-xs"><use href="#icon-tag"></use></svg>NEW</span>' : ''}
        </div>
        
        <p class="vl-video-description">${escapeHtml(video.Video_Description)}</p>

        <div class="vl-alt-stats">
          <span><svg class="vl-icon vl-icon-sm"><use href="#icon-eye"></use></svg>${stats.watchCount}</span>
          <button class="vl-like-btn ${stats.userLiked ? 'active' : ''}" onclick="handleLike('${video.Video_ID}')">
            <svg class="vl-icon vl-icon-sm"><use href="#icon-heart"></use></svg>
            <span>${stats.likeCount}</span>
          </button>
        </div>
        <div class="vl-video-footer" style="margin-top: var(--spacing-sm);">
          ${renderFooterActions(video)}
        </div>
      </div>
    </div>
  `;
}

function renderFooterActions(video, includeShare = true) {
  return `
    ${video.attestationStatus.status !== 'success' ? `
      <button class="vl-btn vl-btn-sm vl-btn-primary" onclick="openAttestationModal('${video.Video_ID}')">
        <svg class="vl-icon vl-icon-sm"><use href="#icon-calendar"></use></svg>
        Extend Attestation
      </button>
    ` : ''}
    <div class="vl-video-actions">
      ${includeShare ? `
        <button class="vl-icon-btn" onclick="shareVideo('${video.Video_ID}')" title="Share video link">
          <svg class="vl-icon vl-icon-sm"><use href="#icon-share"></use></svg>
        </button>
      ` : ''}
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
  `;
}

function getThumbnailUrl(video) {
  if (video.Video_Thumbnail && video.Video_Thumbnail.trim()) {
    return video.Video_Thumbnail;
  }
  return DEFAULT_THUMBNAIL_URL;
}

function renderFocusActions(video, stats) {
  return `
    <button class="vl-like-btn ${stats.userLiked ? 'active' : ''}" onclick="handleLike('${video.Video_ID}')">
      <svg class="vl-icon vl-icon-sm"><use href="#icon-heart"></use></svg>
      <span>${stats.likeCount}</span>
    </button>
    <span class="vl-video-meta-item"><svg class="vl-icon vl-icon-sm"><use href="#icon-eye"></use></svg>${stats.watchCount}</span>
    <button class="vl-icon-btn" onclick="shareVideo('${video.Video_ID}')" title="Share video link">
      <svg class="vl-icon vl-icon-sm"><use href="#icon-share"></use></svg>
    </button>
    ${renderFooterActions(video, false)}
  `;
}

function renderAttestationBadge(video) {
  if (video.attestationStatus.status === 'critical') {
    return `
      <span class="vl-badge vl-badge-danger">
        <svg class="vl-icon vl-icon-sm"><use href="#icon-warning"></use></svg>
        ${video.attestationStatus.message}
      </span>
    `;
  }
  if (video.attestationStatus.status === 'warning') {
    return `
      <span class="vl-badge vl-badge-warning">
        <svg class="vl-icon vl-icon-sm"><use href="#icon-clock"></use></svg>
        ${video.attestationStatus.message}
      </span>
    `;
  }
  return '';
}

function getVideoStats(videoId) {
  const likeSet = videoStats.likes[videoId] || new Set();
  const watches = videoStats.watches[videoId] || 0;
  const userLiked = currentUserName && likeSet.has(currentUserName);
  return { likeCount: likeSet.size, watchCount: watches, userLiked };
}

async function handleLike(videoId) {
  if (!currentUserName) {
    await loadCurrentUser();
  }
  const user = currentUserName || 'Guest User';
  const likeSet = videoStats.likes[videoId] || new Set();
  
  if (likeSet.has(user)) {
    // Unlike: remove record(s) for this user/video/action
    videoStatsRecords = videoStatsRecords.filter(
      r => !(r.Video_ID === videoId && r.User_Name === user && r.Action === 'Like')
    );
    likeSet.delete(user);
    videoStats.likes[videoId] = likeSet;
    
    // Update UI immediately
    if (focusVideoId && focusVideoId === videoId) {
      // Update focus view
      updateFocusActions(videoId);
    } else {
      displayVideos();
    }
    showToast('info', 'Like removed');
    
    // Save to Excel in background
    try {
      await saveVideoStats();
      console.log('Like removed and saved to Excel');
    } catch (err) {
      console.error('Error saving like to Excel:', err);
      showToast('warning', 'Like removed but failed to save. It will sync on next action.');
    }
    return;
  }
  
  const timestamp = new Date().toISOString();
  const record = {
    Log_ID: `${videoId}-${timestamp.replace(/[:.]/g, '-')}`,
    Timestamp: timestamp,
    Video_ID: videoId,
    User_Name: user,
    Action: 'Like'
  };
  
  videoStatsRecords.push(record);
  likeSet.add(user);
  videoStats.likes[videoId] = likeSet;
  
  // Update UI immediately
  if (focusVideoId && focusVideoId === videoId) {
    // Update focus view
    updateFocusActions(videoId);
  } else {
    displayVideos();
  }
  showToast('success', 'Liked');
  
  // Save to Excel in background
  try {
    await saveVideoStats();
    console.log('Like saved to Excel');
  } catch (err) {
    console.error('Error saving like to Excel:', err);
    showToast('warning', 'Liked but failed to save. It will sync on next action.');
  }
}

async function handleWatch(videoId) {
  if (!currentUserName) {
    await loadCurrentUser();
  }
  const user = currentUserName || 'Guest User';
  const timestamp = new Date().toISOString();
  const record = {
    Log_ID: `${videoId}-${timestamp.replace(/[:.]/g, '-')}`,
    Timestamp: timestamp,
    Video_ID: videoId,
    User_Name: user,
    Action: 'Watch'
  };
  
  videoStatsRecords.push(record);
  videoStats.watches[videoId] = (videoStats.watches[videoId] || 0) + 1;
  
  // Save watch to Excel
  try {
    await saveVideoStats();
    console.log('Watch saved to Excel');
  } catch (err) {
    console.error('Error saving watch to Excel:', err);
  }
  
  setFocusVideo(videoId);
}

function toggleInspirationView() {
  inspirationView = !inspirationView;
  const btn = document.getElementById('cardStyleToggle');
  btn.classList.toggle('active', inspirationView);
  displayVideos();
}

function setFocusVideo(videoId) {
  focusVideoId = videoId;
  const container = document.getElementById('focusView');
  const player = document.getElementById('focusPlayer');
  const meta = document.getElementById('focusMeta');
  const titleEl = document.getElementById('focusTitle');
  const stats = getVideoStats(videoId);
  const actionsEl = document.getElementById('focusActions');
  
  const video = allVideos.find(v => v.Video_ID === videoId);
  if (!video || !player || !container) return;
  
  const avatarUrl = userAvatarCache[video.Video_Author_Username] || '';
  
  player.innerHTML = video.Video_Embed_Code || '';
  titleEl.textContent = video.Video_Title || 'Now Playing';
  
  meta.innerHTML = `
    <div class="vl-video-meta-item">
      ${avatarUrl ? `<img src="${avatarUrl}" alt="${escapeHtml(video.Video_Author)}" class="vl-user-avatar" />` : '<svg class="vl-icon"><use href="#icon-user"></use></svg>'}
      <span><strong>${escapeHtml(video.Video_Author || '')}</strong></span>
    </div>
    <div class="vl-video-meta-item">
      <svg class="vl-icon"><use href="#icon-calendar"></use></svg>
      <span>${formatDisplayDate(video.Video_Date)}</span>
    </div>
    ${video.Video_Duration ? `<div class="vl-video-meta-item"><svg class="vl-icon"><use href="#icon-clock"></use></svg>${video.Video_Duration}</div>` : ''}
  `;
  
  // Remove any existing description first
  const existingDescription = document.querySelector('.vl-focus-description');
  if (existingDescription) {
    existingDescription.remove();
  }
  
  // Add description after meta
  const descriptionHtml = `<div class="vl-focus-description">${escapeHtml(video.Video_Description || 'No description available.')}</div>`;
  
  if (actionsEl) {
    actionsEl.innerHTML = renderFocusActions(video, stats);
    // Insert description before actions
    actionsEl.insertAdjacentHTML('beforebegin', descriptionHtml);
  }
  
  buildFocusRail(videoId);
  container.style.display = 'grid';
  document.getElementById('videoGrid').style.display = 'none';
  document.getElementById('videoList').style.display = 'none';
  document.getElementById('pagination').style.display = 'none';
  
  // Disable view toggle buttons when in focus mode
  updateViewToggleState(true);
}

function buildFocusRail(activeId) {
  const rail = document.getElementById('focusRail');
  if (!rail) return;
  rail.innerHTML = filteredVideos
    .filter(v => v.Video_ID !== activeId)
    .map(v => {
    const thumb = getThumbnailUrl(v);
    return `
      <div class="vl-focus-rail-item" onclick="setFocusVideo('${v.Video_ID}')">
        <div class="vl-focus-thumb">
          ${thumb ? `<img src="${thumb}" alt="${escapeHtml(v.Video_Title)}"/>` : ''}
          <div class="vl-thumb-overlay"></div>
        </div>
        <div>
          <h4>${escapeHtml(v.Video_Title)}</h4>
          <p>${escapeHtml(v.Video_Description)}</p>
        </div>
      </div>
    `;
  }).join('');
}

function updateFocusActions(videoId) {
  const video = allVideos.find(v => v.Video_ID === videoId);
  if (!video) return;
  
  const stats = getVideoStats(videoId);
  const actionsEl = document.getElementById('focusActions');
  
  if (actionsEl) {
    actionsEl.innerHTML = renderFocusActions(video, stats);
  }
}

function closeFocusView() {
  const container = document.getElementById('focusView');
  if (container) {
    container.style.display = 'none';
  }
  const player = document.getElementById('focusPlayer');
  if (player) {
    player.innerHTML = '';
  }
  
  // Remove description element if it exists
  const existingDesc = document.querySelector('.vl-focus-description');
  if (existingDesc) {
    existingDesc.remove();
  }
  
  focusVideoId = null;
  
  // Re-enable view toggle buttons
  updateViewToggleState(false);
  
  // Always return to grid view when closing focus mode
  currentView = 'grid';
  document.getElementById('gridViewBtn').classList.add('active');
  document.getElementById('listViewBtn').classList.remove('active');
  
  displayVideos();
}

function displayListView(videos) {
  const gridEl = document.getElementById('videoGrid');
  const listEl = document.getElementById('videoList');
  
  gridEl.style.display = 'none';
  listEl.style.display = focusVideoId ? 'none' : 'block';
  
  // Fetch avatars for all videos before rendering
  const avatarPromises = videos.map(video => {
    if (video.Video_Author_Username) {
      return fetchUserAvatar(video.Video_Author_Username);
    }
    return Promise.resolve();
  });
  
  // Wait a bit for avatars to load, then render
  Promise.all(avatarPromises).finally(() => {
    listEl.innerHTML = videos.map(video => {
      const avatarUrl = userAvatarCache[video.Video_Author_Username] || '';
      const categoryStyle = getCategoryStyle(video.Video_Category);
      const thumb = getThumbnailUrl(video);
      
      return `
        <div class="vl-list-item">
          <div class="vl-list-thumbnail" onclick="handleWatch('${video.Video_ID}'); event.stopPropagation();">
            ${thumb ? `<img src="${thumb}" alt="${escapeHtml(video.Video_Title)}" />` : ''}
            <div class="vl-card-overlay"></div>
          </div>
        
        <div class="vl-list-content">
          <h3 class="vl-list-title">${escapeHtml(video.Video_Title)}</h3>
          <p class="vl-list-description">${escapeHtml(video.Video_Description)}</p>
          
          <div class="vl-list-meta">
            <div class="vl-video-meta-item">
              ${avatarUrl ? `<img src="${avatarUrl}" alt="${escapeHtml(video.Video_Author)}" class="vl-user-avatar" />` : '<svg class="vl-icon vl-icon-sm"><use href="#icon-user"></use></svg>'}
              <span>${escapeHtml(video.Video_Author)}</span>
            </div>
            <div class="vl-video-meta-item">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-calendar"></use></svg>
              <span>${formatDisplayDate(video.Video_Date)}</span>
            </div>
            <span class="vl-badge vl-badge-category" style="background: ${categoryStyle.backgroundColor}; color: ${categoryStyle.color}; border: 1px solid ${categoryStyle.borderColor};">
              ${escapeHtml(video.Video_Category)}
            </span>
            ${video.Video_Duration ? `<span class="vl-video-meta-item"><svg class="vl-icon vl-icon-sm"><use href="#icon-clock"></use></svg>${video.Video_Duration}</span>` : ''}
            ${video.attestationStatus.status === 'critical' ? `
              <span class="vl-badge vl-badge-danger">
                <svg class="vl-icon vl-icon-sm"><use href="#icon-warning"></use></svg>
                ${video.attestationStatus.message}
              </span>
            ` : video.attestationStatus.status === 'warning' ? `
              <span class="vl-badge vl-badge-warning">
                <svg class="vl-icon vl-icon-sm"><use href="#icon-clock"></use></svg>
                ${video.attestationStatus.message}
              </span>
            ` : ''}
          </div>
        </div>
        
        <div class="vl-list-actions">
          ${video.attestationStatus.status !== 'success' ? `
            <button class="vl-btn vl-btn-sm vl-btn-primary" onclick="openAttestationModal('${video.Video_ID}')">
              <svg class="vl-icon vl-icon-sm"><use href="#icon-calendar"></use></svg>
              Extend
            </button>
          ` : ''}
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
  });
}

// ==========================================
// View Toggle
// ==========================================

function updatePagination(totalPages) {
  const paginationEl = document.getElementById('pagination');
  const pagesEl = document.getElementById('paginationPages');
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  
  // Hide pagination in focus mode or if only one page
  if (focusVideoId || totalPages <= 1) {
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
  // Don't allow view change when in focus mode
  if (focusVideoId) {
    showToast('info', 'Close the video player to change view');
    return;
  }
  
  currentView = view;
  
  // Update buttons
  document.getElementById('gridViewBtn').classList.toggle('active', view === 'grid');
  document.getElementById('listViewBtn').classList.toggle('active', view === 'list');
  
  // Save preference
  localStorage.setItem(STORAGE_VIEW_MODE, view);
  
  // Redisplay
  displayVideos();
}

function updateViewToggleState(disabled) {
  const gridBtn = document.getElementById('gridViewBtn');
  const listBtn = document.getElementById('listViewBtn');
  
  if (disabled) {
    gridBtn.style.opacity = '0.5';
    gridBtn.style.pointerEvents = 'none';
    listBtn.style.opacity = '0.5';
    listBtn.style.pointerEvents = 'none';
  } else {
    gridBtn.style.opacity = '1';
    gridBtn.style.pointerEvents = 'auto';
    listBtn.style.opacity = '1';
    listBtn.style.pointerEvents = 'auto';
  }
}

// ==========================================
// Modes (Edit/Delete)
// ==========================================

function toggleEditMode() {
  editMode = !editMode;
  deleteMode = false; // Only one mode at a time
  updateModeButtons();
  
  // Only update display if not in focus mode
  if (!focusVideoId) {
    displayVideos();
  } else {
    // Update focus view to show/hide edit buttons
    updateFocusActions(focusVideoId);
  }
  
  showToast('info', editMode ? 'Edit mode enabled' : 'Edit mode disabled');
}

function toggleDeleteMode() {
  deleteMode = !deleteMode;
  editMode = false; // Only one mode at a time
  updateModeButtons();
  
  // Only update display if not in focus mode
  if (!focusVideoId) {
    displayVideos();
  } else {
    // Update focus view to show/hide delete buttons
    updateFocusActions(focusVideoId);
  }
  
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
  document.getElementById('videoAuthorUsername').value = video.Video_Author_Username || '';
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
    const isEdit = !!videoData.Video_ID && allVideos.some(v => v.Video_ID === videoData.Video_ID);
    
    console.log('=== SUBMIT START ===');
    console.log('Is Edit Mode:', isEdit);
    console.log('Video Data:', videoData);
    console.log('Current allVideos count:', allVideos.length);
    
    if (isEdit) {
      await updateVideo(videoData);
    } else {
      await createVideo(videoData);
    }
    
    console.log('=== SUBMIT COMPLETE ===');
    console.log('Final allVideos count:', allVideos.length);
    
    // Wait a bit for Confluence to process the upload
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // Reload videos to verify save
    console.log('Reloading videos from Confluence...');
    await loadVideos();
    
    // Close modal
    document.getElementById('videoModal').style.display = 'none';
    
    // Show success
    showToast('success', isEdit ? 'Video updated successfully' : 'Video added successfully');
  } catch (error) {
    console.error('=== ERROR SAVING VIDEO ===');
    console.error('Error:', error);
    console.error('Error stack:', error.stack);
    showToast('error', `Failed to save video: ${error.message}`);
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
    Video_Author_Username: document.getElementById('videoAuthorUsername').value.trim(),
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
  
  console.log('Creating video. Total videos before save:', allVideos.length);
  console.log('New video data:', videoData);
  
  // Save to CSV
  await saveVideosToCSV();
  
  // Log audit
  await logAudit('Created', videoData.Video_ID, videoData.Video_Title, null);
}

async function updateVideo(videoData) {
  const index = allVideos.findIndex(v => v.Video_ID === videoData.Video_ID);
  if (index === -1) {
    console.error('Video not found for update:', videoData.Video_ID);
    return;
  }
  
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
  
  console.log('Updating video. Total videos before save:', allVideos.length);
  console.log('Updated video data:', videoData);
  
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
// Attestation Extension
// ==========================================

function openAttestationManager(videoId = null) {
  buildAttestationTable();
  document.getElementById('attestationModal').style.display = 'flex';
  if (videoId) {
    startAttestationExtend(videoId);
  } else {
    resetAttestationForm();
  }
}

function openAttestationModal(videoId) {
  openAttestationManager(videoId);
}

function buildAttestationTable() {
  const body = document.getElementById('attestationTableBody');
  body.innerHTML = allVideos
    .map(video => {
      const badge = renderAttestationBadge(video);
      return `
        <div class="vl-table-row">
          <div>${escapeHtml(video.Video_Title)}</div>
          <div>${formatDisplayDate(video.Video_Attestation)}</div>
          <div>${badge || '<span class="vl-badge vl-badge-attestation">Up to date</span>'}</div>
          <div>
            <button class="vl-btn vl-btn-sm vl-btn-primary" onclick="startAttestationExtend('${video.Video_ID}')">
              Extend
            </button>
          </div>
        </div>
      `;
    })
    .join('');
}

function startAttestationExtend(videoId) {
  const video = allVideos.find(v => v.Video_ID === videoId);
  if (!video) return;
  document.getElementById('currentAttestationDate').value = formatDisplayDate(video.Video_Attestation);
  document.getElementById('newAttestationDate').value = '';
  document.getElementById('extensionReason').value = '';
  document.getElementById('attestationVideoId').value = videoId;
  document.getElementById('attestationFormPanel').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function resetAttestationForm() {
  document.getElementById('currentAttestationDate').value = '';
  document.getElementById('newAttestationDate').value = '';
  document.getElementById('extensionReason').value = '';
  document.getElementById('attestationVideoId').value = '';
}

function closeAttestationModal() {
  document.getElementById('attestationModal').style.display = 'none';
}

async function saveAttestationExtension() {
  const videoId = document.getElementById('attestationVideoId').value;
  const newDate = document.getElementById('newAttestationDate').value;
  const reason = document.getElementById('extensionReason').value;
  
  if (!newDate || !videoId) {
    showToast('warning', 'Pick a video and a new attestation date');
    return;
  }
  
  const video = allVideos.find(v => v.Video_ID === videoId);
  if (!video) return;
  
  const oldDate = video.Video_Attestation;
  
  const saveBtn = document.getElementById('attestationSave');
  const saveText = document.getElementById('attestationSaveText');
  const saveSpinner = document.getElementById('attestationSpinner');
  
  // Show loading
  saveBtn.disabled = true;
  saveText.style.display = 'none';
  saveSpinner.style.display = 'inline-block';
  
  try {
    // Update video
    video.Video_Attestation = newDate;
    
    // Recalculate attestation status
    video.attestationStatus = calculateAttestationStatus(newDate);
    
    // Save videos to CSV
    await saveVideosToCSV();
    
    // Log to attestation_log.csv
    await logAttestationExtension(videoId, video.Video_Title, oldDate, newDate, reason);
    
    // Reload and close
    await loadVideos();
    closeAttestationModal();
    
    showToast('success', 'Attestation date extended successfully');
  } catch (error) {
    console.error('Error extending attestation:', error);
    showToast('error', 'Failed to extend attestation date. Please try again.');
  } finally {
    // Hide loading
    saveBtn.disabled = false;
    saveText.style.display = 'inline';
    saveSpinner.style.display = 'none';
  }
}

async function logAttestationExtension(videoId, videoTitle, oldDate, newDate, reason) {
  try {
    let attestationLog = [];
    try {
      const csvData = await fetchCSVFromConfluence('attestation_log.csv');
      if (csvData) {
        attestationLog = parseCSV(csvData);
      }
    } catch (error) {
      console.log('No existing attestation log, creating new one');
    }
    
    const timestamp = new Date().toISOString();
    const logEntry = {
      Log_ID: timestamp.replace(/[:.]/g, '-'),
      Timestamp: timestamp,
      User_Name: await getConfluenceUser(),
      Video_ID: videoId,
      Video_Title: videoTitle,
      Old_Attestation_Date: oldDate,
      New_Attestation_Date: newDate,
      Reason: reason || 'No reason provided',
      Action: 'Extended'
    };
    
    attestationLog.push(logEntry);
    
    // Convert to CSV
    const headers = Object.keys(logEntry);
    const csvLines = [headers.join(',')];
    attestationLog.forEach(log => {
      const row = headers.map(header => {
        let value = log[header] || '';
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          value = `"${value.replace(/"/g, '""')}"`;
        }
        return value;
      });
      csvLines.push(row.join(','));
    });
    
    const csvContent = csvLines.join('\n');
    await uploadCSVToConfluence('attestation_log.csv', csvContent);
  } catch (error) {
    console.error('Error logging attestation:', error);
    // Don't throw - log failure shouldn't break main functionality
  }
}

// ==========================================
// Stats (Likes / Watches)
// ==========================================

async function loadVideoStats() {
  try {
    const csvData = await fetchCSVFromConfluence(CSV_FILENAME_STATS);
    videoStatsRecords = csvData ? parseCSV(csvData) : [];
  } catch (error) {
    console.log('No existing video stats log, creating new one');
    videoStatsRecords = [];
  }
  
  videoStats = { likes: {}, watches: {} };
  
  videoStatsRecords.forEach(entry => {
    const vid = entry.Video_ID;
    if (entry.Action === 'Like') {
      if (!videoStats.likes[vid]) videoStats.likes[vid] = new Set();
      videoStats.likes[vid].add(entry.User_Name);
    } else if (entry.Action === 'Watch') {
      videoStats.watches[vid] = (videoStats.watches[vid] || 0) + 1;
    }
  });
}

async function saveVideoStats() {
  const headers = ['Log_ID', 'Timestamp', 'Video_ID', 'User_Name', 'Action'];
  const csvLines = [headers.join(',')];
  
  videoStatsRecords.forEach(entry => {
    const row = headers.map(header => {
      let value = entry[header] || '';
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        value = `"${value.replace(/"/g, '""')}"`;
      }
      return value;
    });
    csvLines.push(row.join(','));
  });
  
  const csvContent = csvLines.join('\n');
  await uploadCSVToConfluence(CSV_FILENAME_STATS, csvContent);
}

// ==========================================
// CSV Operations
// ==========================================

async function saveVideosToCSV() {
  console.log('=== BEFORE CSV GENERATION ===');
  console.log('allVideos array length:', allVideos.length);
  console.log('allVideos content:', JSON.stringify(allVideos.map(v => ({
    id: v.Video_ID,
    title: v.Video_Title,
    author: v.Video_Author
  })), null, 2));
  
  const headers = [
    'Video_ID', 'Video_Title', 'Video_Description', 'Video_Author', 'Video_Author_Username',
    'Video_Date', 'Video_Category', 'Video_Document', 'Video_Attestation',
    'Video_Embed_Code', 'Video_Duration', 'Video_Thumbnail', 'Video_Status'
  ];
  
  const csvLines = [headers.join(',')];
  
  allVideos.forEach((video, index) => {
    const row = headers.map(header => {
      let value = video[header] || '';
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        value = `"${value.replace(/"/g, '""')}"`;
      }
      return value;
    });
    csvLines.push(row.join(','));
    
    if (index < 3) {
      console.log(`Row ${index + 1}:`, row.slice(0, 4)); // Log first 4 fields of first 3 rows
    }
  });
  
  const csvContent = csvLines.join('\n');
  
  console.log('=== CSV CONTENT ===');
  console.log('Total lines:', csvLines.length);
  console.log('CSV content length:', csvContent.length);
  console.log('First 500 chars:', csvContent.substring(0, 500));
  console.log('Last 200 chars:', csvContent.substring(csvContent.length - 200));
  
  // Upload to Confluence
  const result = await uploadCSVToConfluence(CSV_FILENAME_VIDEOS, csvContent);
  
  console.log('=== UPLOAD RESULT ===');
  console.log('Result:', result);
  console.log('CSV saved successfully');
  
  return result;
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
    formData.append('comment', `Updated via Video Library App - ${new Date().toISOString()}`);
    
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
    console.error('Error uploading CSV:', error);
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    throw error;  // Re-throw so caller knows it failed
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
    const response = await fetch(`${CONFLUENCE_BASE_URL}/rest/api/user/current`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      },
      credentials: 'include'
    });
    
    if (response.ok) {
      const userData = await response.json();
      return userData.displayName || userData.username || 'Unknown User';
    } else {
      console.log('Could not fetch current user');
      return 'Guest User';
    }
  } catch (error) {
    console.error('Error fetching current user:', error);
    return 'Guest User';
  }
}

async function loadCurrentUser() {
  currentUserName = await getConfluenceUser();
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
    'Video_ID', 'Video_Title', 'Video_Description', 'Video_Author', 'Video_Author_Username',
    'Video_Date', 'Video_Category', 'Video_Document', 'Video_Attestation',
    'Video_Duration', 'Video_Thumbnail', 'Video_Status'
  ];
  
  const csvLines = [headers.join(',')];
  
  videos.forEach(video => {
    const row = headers.map(header => {
      let value = video[header] || '';
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        value = `"${value.replace(/"/g, '""')}"`;
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
  const authorUsernameInput = document.getElementById('videoAuthorUsername');
  const categoryInput = document.getElementById('videoCategory');
  const embedInput = document.getElementById('videoEmbed');
  
  if (!validateField(titleInput, 'titleError', 'Title is required')) isValid = false;
  if (!validateField(descInput, 'descError', 'Description is required')) isValid = false;
  if (!validateField(authorInput, 'authorError', 'Author name is required')) isValid = false;
  if (!validateField(authorUsernameInput, 'authorUsernameError', 'Author username is required')) isValid = false;
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
  
  const titleMap = {
    success: 'Success!',
    error: 'Error!',
    warning: 'Warning!',
    info: 'Information'
  };
  
  const toast = document.createElement('div');
  toast.className = `vl-toast vl-toast-${type}`;
  toast.innerHTML = `
    <div class="vl-toast-icon">
      <svg class="vl-icon vl-icon-md"><use href="#icon-${iconMap[type]}"></use></svg>
    </div>
    <div class="vl-toast-content">
      <h4 class="vl-toast-title">${titleMap[type]}</h4>
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
  
  // Clear existing options (except "All Categories" in filter)
  categoryFilter.innerHTML = '<option value="">All Categories</option>';
  categorySelect.innerHTML = '<option value="">Select a category</option>';
  
  // Populate with dynamic categories
  dynamicCategories.forEach(category => {
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

// Fetch user avatar from Confluence
async function fetchUserAvatar(username) {
  if (!username) {
    return null;
  }
  
  // Return cached value if exists
  if (userAvatarCache[username] !== undefined) {
    return userAvatarCache[username];
  }
  
  try {
    const response = await fetch(`${CONFLUENCE_BASE_URL}/rest/api/user?username=${encodeURIComponent(username)}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      },
      credentials: 'include'
    });
    
    if (response.ok) {
      const userData = await response.json();
      if (userData.profilePicture && userData.profilePicture.path) {
        // Remove /confluence if it's at the start of the path since base URL already has it
        let avatarPath = userData.profilePicture.path;
        if (avatarPath.startsWith('/confluence')) {
          avatarPath = avatarPath.substring(11); // Remove '/confluence'
        }
        const avatarUrl = `${CONFLUENCE_BASE_URL}${avatarPath}`;
        userAvatarCache[username] = avatarUrl;
        return avatarUrl;
      }
    }
  } catch (error) {
    console.error('Error fetching user avatar:', error);
  }
  
  userAvatarCache[username] = null;
  return null;
}

function loadPreferences() {
  // Always default to grid view
  currentView = 'grid';
  document.getElementById('gridViewBtn').classList.add('active');
  document.getElementById('listViewBtn').classList.remove('active');
}

// ==========================================
// Share Video Function
// ==========================================

function shareVideo(videoId) {
  const video = allVideos.find(v => v.Video_ID === videoId);
  if (!video) return;
  
  // Create shareable URL with video ID parameter
  const baseUrl = window.location.href.split('?')[0];
  const shareUrl = `${baseUrl}?video=${videoId}`;
  
  // Copy to clipboard
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(shareUrl)
      .then(() => {
        showToast('success', `Share link copied! Anyone with this link can watch "${video.Video_Title}"`);
      })
      .catch(err => {
        console.error('Failed to copy to clipboard:', err);
        // Fallback: show the URL in a prompt
        prompt('Copy this link to share:', shareUrl);
      });
  } else {
    // Fallback for older browsers
    prompt('Copy this link to share:', shareUrl);
  }
}
