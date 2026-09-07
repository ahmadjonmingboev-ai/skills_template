document.getElementById('videotrainings').innerHTML = `
<div class="icgds" id="videoLibraryApp">

  <!-- ---- TOOLBAR - search right next to refresh ---------------------------------------------------------------- -->
  <div class="vl-toolbar">
    <div class="vl-toolbar-left">
      <h1 class="vl-title">Video Materials Library</h1>
    </div>
    <div class="vl-toolbar-right">

      <!-- Search sits directly before refresh -->
      <div class="vl-search-container">
        <i class="lmnicon lmnicon-search lmnicon-sm vl-search-icon" aria-hidden="true"></i>
        <input type="text" class="lmn-input vl-search-input" id="searchInput"
          placeholder="Search videos..." aria-label="Search videos" />
        <button class="vl-search-clear" id="searchClear" aria-label="Clear search" style="display:none;">
          <i class="lmnicon lmnicon-close lmnicon-sm"></i>
        </button>
      </div>

      <button class="vl-refresh-btn" id="refreshBtn" title="Refresh">
        <i class="lmnicon lmnicon-reload lmnicon-sm"></i>
      </button>

      <div class="vl-view-toggle">
        <button class="vl-view-btn active" id="gridViewBtn" title="Grid view">
          <i class="lmnicon lmnicon-layout-grid lmnicon-sm"></i>
        </button>
        <button class="vl-view-btn" id="listViewBtn" title="List view">
          <i class="lmnicon lmnicon-layout-rows lmnicon-sm"></i>
        </button>
      </div>

      <div class="vl-config-dropdown">
        <button class="vl-config-btn" id="configBtn" title="Settings">
          <i class="lmnicon lmnicon-setting lmnicon-sm"></i>
        </button>
        <div class="vl-dropdown-menu" id="configMenu">
          <button class="vl-dropdown-item" id="addVideoBtn">
            <i class="lmnicon lmnicon-add lmnicon-sm"></i>
            <span>Add New Video</span>
          </button>
          <button class="vl-dropdown-item" id="editModeBtn">
            <i class="lmnicon lmnicon-edit lmnicon-sm"></i>
            <span>Edit Mode</span>
          </button>
          <button class="vl-dropdown-item" id="deleteModeBtn">
            <i class="lmnicon lmnicon-trash lmnicon-sm"></i>
            <span>Delete Mode</span>
          </button>
          <button class="vl-dropdown-item" id="attestationManagerBtn">
            <i class="lmnicon lmnicon-calendar-dots lmnicon-sm"></i>
            <span>Attestation Manager</span>
          </button>
          <div class="vl-dropdown-divider"></div>
          <button class="vl-dropdown-item" id="exportVideosBtn">
            <i class="lmnicon lmnicon-download lmnicon-sm"></i>
            <span>Export Videos</span>
          </button>
          <button class="vl-dropdown-item" id="exportAuditBtn">
            <i class="lmnicon lmnicon-download lmnicon-sm"></i>
            <span>Export Audit Log</span>
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ---- FILTER BAR - filters left, results + pagination right ---- -->
  <div class="vl-filter-bar">
    <div class="vl-filter-left">
      <div class="vl-filter-group">
        <i class="lmnicon lmnicon-filter-alt lmnicon-sm" style="color:var(--text_weak,#46637F)"></i>
        <select class="vl-select" id="categoryFilter" aria-label="Filter by category">
          <option value="">All Categories</option>
        </select>
      </div>
      <div class="vl-filter-group">
        <i class="lmnicon lmnicon-sort lmnicon-sm" style="color:var(--text_weak,#46637F)"></i>
        <select class="vl-select" id="sortSelect" aria-label="Sort videos">
          <option value="newest">Newest First</option>
          <option value="oldest">Oldest First</option>
          <option value="title-asc">Title (A-Z)</option>
          <option value="title-desc">Title (Z-A)</option>
          <option value="attestation">Attestation (Urgent First)</option>
        </select>
      </div>
      <button class="lmn-btn lmn-btn-standard vl-filter-clear-btn" id="clearFiltersBtn" style="display:none;">
        <i class="lmnicon lmnicon-close lmnicon-sm"></i>
        Clear
      </button>
    </div>
    <div class="vl-filter-right">
      <span class="vl-results-count" id="resultsCount">Showing 0 videos</span>
      <!-- Inline pagination - built by JS -->
      <div class="vl-pagination-inline" id="paginationInline" style="display:none;"></div>
    </div>
  </div>

  <!-- ---- LOADING SKELETON ---------------------------------------------------------------- -->
  <div class="vl-skeleton-container" id="skeletonContainer">
    <div class="vl-loading-overlay">
      <div class="lmn-loading">
        <svg class="lmn-loading-icon lmn-loading-svg" viewBox="25 25 50 50" aria-hidden="true">
          <circle class="lmn-loading-svg-path" cx="50" cy="50" r="20" fill="none" stroke-width="4"/>
        </svg>
        <p class="lmn-loading-text">Loading videos...</p>
      </div>
    </div>
    <div class="vl-skeleton-card"></div>
    <div class="vl-skeleton-card"></div>
    <div class="vl-skeleton-card"></div>
    <div class="vl-skeleton-card"></div>
    <div class="vl-skeleton-card"></div>
  </div>

  <!-- ---- FOCUS / PLAYER VIEW ---------------------------------------------------------------- -->
  <div class="vl-focus-view" id="focusView" style="display:none;">
    <div class="vl-focus-main">
      <div class="vl-focus-header">
        <h3 id="focusTitle">Now Playing</h3>
        <button class="lmn-btn lmn-btn-standard vl-focus-close-btn" id="focusCloseBtn">
          <i class="lmnicon lmnicon-close lmnicon-sm"></i>
          Close
        </button>
      </div>
      <div class="vl-focus-player" id="focusPlayer"></div>
      <div class="vl-focus-meta" id="focusMeta"></div>
      <div class="vl-focus-actions" id="focusActions"></div>
    </div>
    <div class="vl-focus-rail">
      <div class="vl-focus-rail-title">More Videos</div>
      <div class="vl-focus-rail-list" id="focusRail"></div>
    </div>
  </div>

  <!-- ---- GRID VIEW ---------------------------------------------------------------- -->
  <div class="vl-video-grid" id="videoGrid"></div>

  <!-- ---- LIST VIEW ---------------------------------------------------------------- -->
  <div class="vl-video-list" id="videoList" style="display:none;"></div>

  <!-- ---- EMPTY STATE ---------------------------------------------------------------- -->
  <div class="vl-empty-state" id="emptyState" style="display:none;">
    <i class="lmnicon lmnicon-blank-document vl-empty-icon-lmn" aria-hidden="true"></i>
    <h3 class="vl-empty-title" id="emptyTitle">No videos found</h3>
    <p class="vl-empty-text" id="emptyText">Try adjusting your filters or search terms</p>
    <button class="lmn-btn lmn-btn-primary" id="emptyActionBtn" style="display:none;">
      <i class="lmnicon lmnicon-add lmnicon-sm"></i>
      Add Video
    </button>
  </div>

  <!-- ---- ADD / EDIT VIDEO MODAL ---------------------------------------------------------------- -->
  <div class="vl-modal-overlay" id="videoModal" style="display:none;">
    <div class="lmn-modal vl-modal-large">
      <div class="lmn-modal-content">
        <div class="lmn-modal-header">
          <span class="lmn-modal-title" id="modalTitle">Add New Video</span>
          <button class="lmn-action-icon" id="modalClose" aria-label="Close">
            <i class="lmnicon lmnicon-close lmnicon-md"></i>
          </button>
        </div>
        <div class="lmn-modal-body">
          <form id="videoForm" novalidate>

            <div class="vl-form-section-label">Core Details</div>

            <div class="vl-form-row">
              <div class="lmn-form-group">
                <label for="videoTitle" class="lmn-label">Video Title <span class="vl-required">*</span></label>
                <input type="text" id="videoTitle" class="lmn-input" maxlength="100" required />
                <div class="vl-char-counter"><span id="titleCounter">0</span>/100</div>
                <small class="lmn-invalid" id="titleError"></small>
              </div>
            </div>

            <div class="vl-form-row">
              <div class="lmn-form-group">
                <label for="videoDescription" class="lmn-label">Description <span class="vl-required">*</span></label>
                <textarea id="videoDescription" class="lmn-textarea" maxlength="500" rows="3" required></textarea>
                <div class="vl-char-counter"><span id="descCounter">0</span>/500</div>
                <small class="lmn-invalid" id="descError"></small>
              </div>
            </div>

            <div class="vl-form-row vl-form-row-2">
              <div class="lmn-form-group">
                <label for="videoAuthor" class="lmn-label">Author Name <span class="vl-required">*</span></label>
                <input type="text" id="videoAuthor" class="lmn-input" maxlength="50" required />
                <small class="lmn-invalid" id="authorError"></small>
              </div>
              <div class="lmn-form-group">
                <label for="videoAuthorUsername" class="lmn-label">Author Username <span class="vl-required">*</span></label>
                <input type="text" id="videoAuthorUsername" class="lmn-input" maxlength="50" placeholder="e.g., john.doe" required />
                <small class="lmn-hint-text">Confluence username for avatar</small>
                <small class="lmn-invalid" id="authorUsernameError"></small>
              </div>
            </div>

            <div class="vl-form-row vl-form-row-2">
              <div class="lmn-form-group">
                <label for="videoCategory" class="lmn-label">Category <span class="vl-required">*</span></label>
                <select id="videoCategory" class="lmn-input" required>
                  <option value="">Select a category</option>
                </select>
                <small class="lmn-invalid" id="categoryError"></small>
              </div>
              <div class="lmn-form-group">
                <label for="videoDate" class="lmn-label">Publish Date</label>
                <input type="date" id="videoDate" class="lmn-input" />
                <small class="lmn-hint-text">Defaults to today if blank</small>
              </div>
            </div>

            <div class="vl-form-section-label" style="margin-top:8px;">Additional Details</div>

            <div class="vl-form-row vl-form-row-2">
              <div class="lmn-form-group">
                <label for="videoDuration" class="lmn-label">Duration</label>
                <input type="text" id="videoDuration" class="lmn-input" placeholder="MM:SS (e.g., 05:30)" />
                <small class="lmn-invalid" id="durationError"></small>
              </div>
              <div class="lmn-form-group">
                <label for="videoDocument" class="lmn-label">Related Document URL</label>
                <input type="url" id="videoDocument" class="lmn-input" placeholder="https://..." />
                <small class="lmn-invalid" id="documentError"></small>
              </div>
            </div>

            <div class="vl-form-row">
              <div class="lmn-form-group">
                <label for="videoThumbnail" class="lmn-label">Custom Thumbnail URL</label>
                <input type="url" id="videoThumbnail" class="lmn-input" placeholder="https://..." />
                <small class="lmn-hint-text">Optional - leave blank to use default</small>
                <small class="lmn-invalid" id="thumbnailError"></small>
              </div>
            </div>

            <div class="vl-form-section-label" style="margin-top:8px;">Embed Code</div>

            <div class="vl-form-row">
              <div class="lmn-form-group">
                <label for="videoEmbed" class="lmn-label">Video Embed Code <span class="vl-required">*</span></label>
                <textarea id="videoEmbed" class="lmn-textarea" rows="4" placeholder="Paste iframe code from SharePoint..." required></textarea>
                <small class="lmn-hint-text">Paste the complete iframe embed code</small>
                <small class="lmn-invalid" id="embedError"></small>
              </div>
            </div>

            <input type="hidden" id="videoId" />
          </form>
        </div>
        <div class="lmn-modal-footer">
          <button type="button" class="lmn-btn lmn-btn-standard" id="modalCancel">Cancel</button>
          <button type="submit" class="lmn-btn lmn-btn-primary" id="modalSave" form="videoForm">
            <i class="lmnicon lmnicon-check lmnicon-sm"></i>
            <span id="saveText">Save Video</span>
            <i class="lmnicon lmnicon-spinner lmnicon-sm vl-spinner" id="saveSpinner" style="display:none;"></i>
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ---- DELETE MODAL ---------------------------------------------------------------- -->
  <div class="vl-modal-overlay" id="deleteModal" style="display:none;">
    <div class="lmn-modal vl-modal-small">
      <div class="lmn-modal-content">
        <div class="lmn-modal-header">
          <span class="lmn-modal-title">Delete Video</span>
          <button class="lmn-action-icon" id="deleteModalClose" aria-label="Close">
            <i class="lmnicon lmnicon-close lmnicon-md"></i>
          </button>
        </div>
        <div class="lmn-modal-body">
          <div class="lmn-alert lmn-alert-danger lmn-alert-low-contrast lmn-alert-multi-line" role="alert">
            <span class="lmn-alert-icon" aria-hidden="true"></span>
            <div class="lmn-alert-text">
              <p>Are you sure you want to delete <strong id="deleteVideoTitle"></strong>?</p>
              <p style="margin-top:6px;font-weight:700;">This action cannot be undone.</p>
            </div>
          </div>
        </div>
        <div class="lmn-modal-footer">
          <button type="button" class="lmn-btn lmn-btn-standard" id="deleteCancel">Cancel</button>
          <button type="button" class="lmn-btn lmn-btn-danger" id="deleteConfirm">
            <i class="lmnicon lmnicon-trash lmnicon-sm"></i>
            <span id="deleteText">Delete Video</span>
            <i class="lmnicon lmnicon-spinner lmnicon-sm vl-spinner" id="deleteSpinner" style="display:none;"></i>
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ---- ATTESTATION MODAL ---------------------------------------------------------------- -->
  <div class="vl-modal-overlay" id="attestationModal" style="display:none;">
    <div class="lmn-modal vl-modal-large">
      <div class="lmn-modal-content">
        <div class="lmn-modal-header">
          <span class="lmn-modal-title">Attestation Manager</span>
          <button class="lmn-action-icon" id="attestationModalClose" aria-label="Close">
            <i class="lmnicon lmnicon-close lmnicon-md"></i>
          </button>
        </div>
        <div class="lmn-modal-body">
          <div class="vl-attestation-layout">
            <div class="vl-attestation-list">
              <div class="vl-table">
                <div class="vl-table-head">
                  <div>Title</div><div>Attestation</div><div>Status</div><div>Action</div>
                </div>
                <div class="vl-table-body" id="attestationTableBody"></div>
              </div>
            </div>
            <div class="vl-attestation-form" id="attestationFormPanel">
              <div class="vl-form-section-label">Extend Attestation</div>
              <div class="vl-form-row">
                <div class="lmn-form-group">
                  <label class="lmn-label">Current Attestation Date</label>
                  <input type="text" id="currentAttestationDate" class="lmn-input" readonly />
                </div>
              </div>
              <div class="vl-form-row">
                <div class="lmn-form-group">
                  <label class="lmn-label">New Attestation Date <span class="vl-required">*</span></label>
                  <input type="date" id="newAttestationDate" class="lmn-input" required />
                </div>
              </div>
              <div class="vl-form-row">
                <div class="lmn-form-group">
                  <label class="lmn-label">Reason for Extension</label>
                  <textarea id="extensionReason" class="lmn-textarea" rows="3"
                    placeholder="Enter the reason for extending the attestation date..."></textarea>
                </div>
              </div>
              <input type="hidden" id="attestationVideoId" />
            </div>
          </div>
        </div>
        <div class="lmn-modal-footer">
          <button type="button" class="lmn-btn lmn-btn-standard" id="attestationCancel">Cancel</button>
          <button type="button" class="lmn-btn lmn-btn-primary" id="attestationSave">
            <i class="lmnicon lmnicon-calendar-dots lmnicon-sm"></i>
            <span id="attestationSaveText">Extend Attestation</span>
            <i class="lmnicon lmnicon-spinner lmnicon-sm vl-spinner" id="attestationSpinner" style="display:none;"></i>
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ---- TOAST ---------------------------------------------------------------- -->
  <div class="vl-toast-container" id="toastContainer"></div>

</div>
`;
