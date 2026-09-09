document.getElementById('rebrandedhome').innerHTML = `
<div class="main-container">
    <!-- Greeting Section -->
    <div class="greeting-section">
        <div class="logo-container">
            <svg xmlns="http://www.w3.org/2000/svg" width="204" height="118" viewBox="0 0 204 118" fill="none">
             </svg>
        </div>
        <h1 class="greeting-title">
            <span id="greeting">Good morning</span>,
            <span class="name-underline">
                <span id="userName">Team</span>
                <svg viewBox="0 0 140 24" fill="none" preserveAspectRatio="none" aria-hidden="true">
                    <path d="M6 16 Q 70 24, 134 14" stroke="#D97757" stroke-width="3" stroke-linecap="round" fill="none" />
                </svg>
            </span>
        </h1>
        <p class="greeting-subtitle">No more searching across multiple sources. Your complete Investor Services Training & Development platform is right here.</p>
    </div>

    <!-- Search Section -->
    <div class="search-section">
        <div class="search-container">
            <div class="search-box" id="searchBox">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 512 512" class="search-icon">
                    <path d="M497.913 497.913c-18.782 18.782-49.225 18.782-68.008 0l-84.862-84.863c-34.889 22.382-76.13 35.717-120.659 35.717C100.469 448.767 0 348.312 0 224.383S100.469 0 224.384 0c123.931 0 224.384 100.452 224.384 224.383 0 44.514-13.352 85.771-35.718 120.676l84.863 84.863c18.782 18.782 18.782 49.209 0 67.991zM224.384 64.109c-88.511 0-160.274 71.747-160.274 160.273s71.764 160.274 160.274 160.274c88.525 0 160.273-71.748 160.273-160.274S312.909 64.109 224.384 64.109z" fill="currentColor"></path>
                </svg>
                <input type="text" placeholder="Search in Resource Hub..." class="search-input" id="searchInput" autocomplete="off">
                <div class="search-loading hidden" id="searchLoading">
                    <div class="spinner"></div>
                </div>
            </div>

            <!-- Search Results Dropdown -->
            <div class="search-results hidden" id="searchResults">
                <div class="results-list" id="resultsList"></div>
                <div class="no-results hidden" id="noResults">
                    <p>No results found</p>
                </div>
            </div>
        </div>
    </div>

    <!-- Suggestion Buttons -->
    <div class="suggestion-buttons">

        <!-- Video Library: play button / film icon -->
        <button class="suggestion-btn" onclick="window.open('url_placeholder_forlinks', '_blank')">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
                <polygon points="10 8 16 11 10 14 10 8" fill="currentColor" stroke="none"></polygon>
                <line x1="8" y1="21" x2="16" y2="21"></line>
                <line x1="12" y1="17" x2="12" y2="21"></line>
            </svg>
            Video Library
        </button>

        <!-- Links Library: chain/link icon -->
        <button class="suggestion-btn" onclick="window.open('url_placeholder_forlinks', '_blank')">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path>
                <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path>
            </svg>
            Links Library
        </button>

        <!-- Marketplace Library: shopping bag / store icon -->
        <button class="suggestion-btn" onclick="window.open('url_placeholder_forlinks', '_blank')">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"></path>
                <line x1="3" y1="6" x2="21" y2="6"></line>
                <path d="M16 10a4 4 0 0 1-8 0"></path>
            </svg>
            Orders Library
        </button>

        <!-- Procedures & Guides: clipboard / checklist icon -->
        <button class="suggestion-btn" onclick="window.open('url_placeholder_forlinks', '_blank')">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"></path>
                <rect x="9" y="3" width="6" height="4" rx="1" ry="1"></rect>
                <line x1="9" y1="12" x2="15" y2="12"></line>
                <line x1="9" y1="16" x2="13" y2="16"></line>
            </svg>
            Procedures & Guides
        </button>

        <!-- New Joiner Resources: user-plus / person with star icon -->
        <button class="suggestion-btn" onclick="window.open('url_placeholder_forlinks', '_blank')">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                <circle cx="8.5" cy="7" r="4"></circle>
                <line x1="20" y1="8" x2="20" y2="14"></line>
                <line x1="23" y1="11" x2="17" y2="11"></line>
            </svg>
            New Joiner Resources
        </button>

    </div>
</div>
`;