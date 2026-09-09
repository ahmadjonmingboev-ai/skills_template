// Set greeting based on time of day
const currentHour = new Date().getHours();
let greeting = 'Good morning';
if (currentHour >= 12 && currentHour < 18) {
    greeting = 'Good afternoon';
} else if (currentHour >= 18) {
    greeting = 'Good evening';
}
document.getElementById('greeting').textContent = greeting;

// Current user object
let currentUser = null;

// Function to get current user from Confluence API
async function getCurrentUser() {
    try {
        const response = await fetch('https://cedt-confluence.nam.nsroot.net/confluence/rest/api/user/current', {
            method: 'GET',
            headers: {
                'Accept': 'application/json'
            },
            credentials: 'include'
        });

        if (!response.ok) {
            throw new Error(`Failed to get current user: ${response.status} ${response.statusText}`);
        }

        const userData = await response.json();

        return {
            username: userData.username || userData.accountId || 'unknown',
            displayName: userData.displayName || userData.username || 'Unknown User',
            type: userData.type || 'unknown'
        };
    } catch (error) {
        console.warn('Error fetching current user:', error);
        return null;
    }
}

// Function to check if user is logged in
async function isUserLoggedIn() {
    try {
        const user = await getCurrentUser();
        if (
            user &&
            user.type !== 'anonymous' &&
            user.displayName !== 'Anonymous' &&
            user.username !== 'unknown' &&
            user.username !== 'anonymous'
        ) {
            currentUser = user;
            return true;
        } else {
            return false;
        }
    } catch (error) {
        console.error('Error checking user login status:', error);
        return false;
    }
}

// Initialize user authentication and update UI
async function initializeUser() {
    const isLoggedIn = await isUserLoggedIn();
    const userNameElement = document.getElementById('userName');

    if (isLoggedIn && currentUser) {
        let firstName = 'Team';

        if (currentUser.displayName.includes(',')) {
            const parts = currentUser.displayName.split(',');
            if (parts.length > 1) {
                const namePart = parts[1].split('[')[0].trim();
                firstName = namePart.split(' ')[0];
            }
        } else {
            firstName = currentUser.displayName.split(' ')[0];
        }

        userNameElement.textContent = firstName;
    } else {
        userNameElement.textContent = 'Team';
    }
}

// Call user initialization on page load
initializeUser();

// Search functionality
const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');
const resultsList = document.getElementById('resultsList');
const noResults = document.getElementById('noResults');
const searchLoading = document.getElementById('searchLoading');

let searchTimeout;
let currentAbortController;

const CONFLUENCE_BASE_URL = 'https://cedt-confluence.nam.nsroot.net/confluence';
const CONFLUENCE_SPACE = 'ischtd'; // Your Confluence space key

// Function to perform search
async function performSearch(query) {
    if (query.length < 2) {
        hideResults();
        return;
    }

    searchLoading.classList.remove('hidden');

    try {
        currentAbortController = new AbortController();

        // Search only in LPTM space, exclude attachments
        const cqlQuery = `(text ~ "*${query}*" AND space = ${CONFLUENCE_SPACE} AND type = page)`;
        const encodedCql = encodeURIComponent(cqlQuery);
        const apiUrl = `${CONFLUENCE_BASE_URL}/rest/api/content/search?cql=${encodedCql}&limit=10&expand=space,body.view`;

        const response = await fetch(apiUrl, {
            method: 'GET',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            signal: currentAbortController.signal,
            credentials: 'include'
        });

        if (!response.ok) {
            throw new Error(`Search failed: ${response.status}`);
        }

        const data = await response.json();
        displayResults(data.results || [], query);

    } catch (error) {
        if (error.name !== 'AbortError') {
            console.error('Search error:', error);
            showNoResults();
        }
    } finally {
        searchLoading.classList.add('hidden');
    }
}

// Function to strip HTML tags
function stripHtmlTags(html) {
    if (!html) return '';

    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;

    const scripts = tempDiv.querySelectorAll('script, style');
    scripts.forEach(script => script.remove());

    let text = tempDiv.textContent || tempDiv.innerText || '';
    text = text.replace(/\s+/g, ' ').trim();

    return text;
}

// Function to display results
function displayResults(results, query) {
    resultsList.innerHTML = '';

    if (!results || results.length === 0) {
        showNoResults();
        return;
    }

    results.forEach(result => {
        const resultElement = createResultElement(result, query);
        resultsList.appendChild(resultElement);
    });

    showResults();
}

// Function to create result element
function createResultElement(result, query) {
    const div = document.createElement('div');
    div.className = 'result-item';

    const pageUrl = result._links && result._links.webui ? `${CONFLUENCE_BASE_URL}${result._links.webui}` : '';

    let excerpt = '';
    if (result.body && result.body.view && result.body.view.value) {
        const cleanText = stripHtmlTags(result.body.view.value);
        excerpt = cleanText.length > 150 ? cleanText.substring(0, 150) + '...' : cleanText;
    } else {
        excerpt = 'Click to view content';
    }

    const highlightedTitle = highlightSearchTerms(result.title, query);
    const highlightedExcerpt = highlightSearchTerms(excerpt, query);

    div.innerHTML = `
        <div class="result-icon">
            <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
            </svg>
        </div>
        <div class="result-content">
            <div class="result-title">${highlightedTitle}</div>
            <div class="result-excerpt">${highlightedExcerpt}</div>
            <div class="result-url">${pageUrl}</div>
        </div>
    `;

    if (pageUrl) {
        div.addEventListener('click', () => {
            window.open(pageUrl, '_blank');
        });
    }

    return div;
}

// Function to highlight search terms
function highlightSearchTerms(text, query) {
    if (!query || !text) return text;
    const regex = new RegExp(`(${query.split(' ').join('|')})`, 'gi');
    return text.replace(regex, '<mark style="background-color: #fef3c7; padding: 0 2px; border-radius: 2px;">$1</mark>');
}

// Function to show results
function showResults() {
    noResults.classList.add('hidden');
    searchResults.classList.remove('hidden');
}

// Function to show no results
function showNoResults() {
    resultsList.innerHTML = '';
    noResults.classList.remove('hidden');
    searchResults.classList.remove('hidden');
}

// Function to hide results
function hideResults() {
    searchResults.classList.add('hidden');
    noResults.classList.add('hidden');
}

// Event listener for search input
searchInput.addEventListener('input', function() {
    const query = this.value.trim();

    if (searchTimeout) {
        clearTimeout(searchTimeout);
    }

    if (currentAbortController) {
        currentAbortController.abort();
    }

    if (query.length < 2) {
        hideResults();
        return;
    }

    searchTimeout = setTimeout(() => {
        performSearch(query);
    }, 300);
});

// Hide results when clicking outside
document.addEventListener('click', function(event) {
    const searchBox = document.getElementById('searchBox');
    if (searchBox && !searchBox.contains(event.target) &&
        searchResults && !searchResults.contains(event.target)) {
        hideResults();
    }
});

// Handle escape key
searchInput.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        hideResults();
        this.blur();
    }
});