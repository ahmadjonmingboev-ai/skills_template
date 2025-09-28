// Reporting Page JavaScript
// Include Chart.js: <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>

// Configuration
const CONFIG = {
    confluence: {
        domain: 'https://www.myconfluence.net',
        pageId: '123345346',
        api: {
            currentUser: '/rest/api/user/current',
            attachment: '/rest/api/content/{pageId}/child/attachment'
        }
    },
    files: {
        userDatabase: 'userdatabase.xlsx',
        mainTaskDatabase: 'maintaskdatabase.xlsx',
        taskMapping: 'taskmapping.xlsx'
    }
};

// API Helper Functions
const API = {
    getCurrentUser: async function() {
        try {
            const response = await fetch(CONFIG.confluence.domain + CONFIG.confluence.api.currentUser, {
                credentials: 'same-origin'
            });
            if (!response.ok) return null;
            const userData = await response.json();
            return {
                displayName: userData.displayName || userData.fullName || userData.name || 'Unknown',
                key: userData.key || userData.username || userData.accountId || userData.name,
                email: userData.emailAddress || userData.email,
                ...userData
            };
        } catch (error) {
            console.warn('Error fetching current user:', error);
            return null;
        }
    },
    
    getAttachments: async function() {
        try {
            const url = CONFIG.confluence.domain +
                       CONFIG.confluence.api.attachment.replace('{pageId}', CONFIG.confluence.pageId);
            const response = await fetch(url, { credentials: 'same-origin' });
            if (!response.ok) throw new Error('Failed to fetch attachments');
            const data = await response.json();
            return data.results || [];
        } catch (error) {
            console.error('Error fetching attachments:', error);
            return [];
        }
    },
    
    downloadAttachment: async function(attachmentUrl) {
        try {
            const url = attachmentUrl.startsWith('http') ? attachmentUrl : CONFIG.confluence.domain + attachmentUrl;
            const response = await fetch(url, {
                credentials: 'same-origin',
                headers: { 'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
            });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            return await response.arrayBuffer();
        } catch (error) {
            console.error('Error downloading attachment:', error);
            return null;
        }
    }
};

// Global variables
let currentUser = null;
let allUsers = [];
let allSkillData = {};
let taskMapping = [];
let availableMarkets = [];
let categories = [];
let attachments = {};
let chartInstances = {};

// Initialize application
async function initializeReporting() {
    showLoading(true, 'Loading reporting data...');
    try {
        // Get current user
        const confluenceUser = await API.getCurrentUser();
        if (!confluenceUser) {
            showToast('Could not get current user. Please ensure you are logged into Confluence.', 'error');
            return;
        }
        
        currentUser = {
            username: confluenceUser.key,
            displayName: confluenceUser.displayName,
            email: confluenceUser.email
        };
        
        // Load attachments
        await loadAttachments();
        
        // Check access
        const hasAccess = await checkReportingAccess();
        if (!hasAccess) {
            showNoAccess();
            return;
        }
        
        // Load all data
        await loadAllData();
        
        // Calculate metrics
        calculateMetrics();
        
        // Render charts
        renderCharts();
        
        // Populate tables
        populateTables();
        
        // Initialize market readiness
        renderMarketReadiness();
        
    } catch (error) {
        console.error('Initialization error:', error);
        showToast('Failed to initialize reporting. Please refresh the page.', 'error');
    } finally {
        showLoading(false);
    }
}

// Load attachments
async function loadAttachments() {
    const attachmentList = await API.getAttachments();
    attachmentList.forEach(att => {
        attachments[att.title] = att._links.download || att._links.webui;
    });
}

// Fetch Excel data
async function fetchExcelData(filename) {
    try {
        const downloadUrl = attachments[filename];
        if (!downloadUrl) throw new Error(`Attachment ${filename} not found`);
        
        const arrayBuffer = await API.downloadAttachment(downloadUrl);
        if (!arrayBuffer) throw new Error(`Failed to download ${filename}`);
        
        return XLSX.read(arrayBuffer, { type: 'array' });
    } catch (error) {
        console.error(`Error fetching ${filename}:`, error);
        throw error;
    }
}

// Check reporting access
async function checkReportingAccess() {
    try {
        const workbook = await fetchExcelData(CONFIG.files.userDatabase);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        allUsers = XLSX.utils.sheet_to_json(worksheet);
        
        const userData = allUsers.find(u => u.username === currentUser.username);
        if (!userData) return false;
        
        currentUser = { ...currentUser, ...userData };
        
        // Admin and managers can access reporting
        return userData.role === 'admin' || userData.role === 'manager';
    } catch (error) {
        console.error('Failed to check access:', error);
        return false;
    }
}

// Load all data
async function loadAllData() {
    showLoadingMessage('Loading user data...');
    
    // Load user database (already loaded in checkAccess)
    
    // Load task mapping
    showLoadingMessage('Loading task mappings...');
    const mappingWorkbook = await fetchExcelData(CONFIG.files.taskMapping);
    const mappingSheet = mappingWorkbook.Sheets[mappingWorkbook.SheetNames[0]];
    taskMapping = XLSX.utils.sheet_to_json(mappingSheet);
    
    // Extract markets and categories
    if (taskMapping.length > 0) {
        const standardColumns = ['Item_ID', 'Category', 'Task_Group', 'Task_Name'];
        availableMarkets = Object.keys(taskMapping[0]).filter(key => !standardColumns.includes(key));
        categories = [...new Set(taskMapping.map(t => t.Category))].filter(Boolean);
    }
    
    // Load main task database with all user skills
    showLoadingMessage('Loading skill data...');
    const mainWorkbook = await fetchExcelData(CONFIG.files.mainTaskDatabase);
    
    // Process each user sheet
    let processedUsers = 0;
    const totalUsers = mainWorkbook.SheetNames.filter(name => name !== 'Template').length;
    
    for (const sheetName of mainWorkbook.SheetNames) {
        if (sheetName === 'Template') continue;
        
        processedUsers++;
        showLoadingMessage(`Processing user ${processedUsers}/${totalUsers}...`);
        
        const userSheet = mainWorkbook.Sheets[sheetName];
        const userData = XLSX.utils.sheet_to_json(userSheet);
        
        // Store user skill data
        allSkillData[sheetName] = userData;
    }
    
    console.log('Loaded data for', Object.keys(allSkillData).length, 'users');
}

// Calculate metrics
function calculateMetrics() {
    // Overall statistics
    let totalSkills = 0;
    let completedSkills = 0;
    
    Object.values(allSkillData).forEach(userSkills => {
        userSkills.forEach(skill => {
            availableMarkets.forEach(market => {
                if (skill[market] !== undefined && skill[market] !== '') {
                    totalSkills++;
                    if (skill[market] === 'yes') {
                        completedSkills++;
                    }
                }
            });
        });
    });
    
    const overallCompletion = totalSkills > 0 ? Math.round((completedSkills / totalSkills) * 100) : 0;
    
    // Update metric cards
    document.getElementById('totalUsersMetric').textContent = Object.keys(allSkillData).length;
    document.getElementById('overallCompletion').textContent = overallCompletion + '%';
    document.getElementById('totalTasksMetric').textContent = taskMapping.length;
    document.getElementById('activeMarkets').textContent = availableMarkets.length;
}

// Render charts
function renderCharts() {
    renderTeamPerformanceChart();
    renderCategoryProgressChart();
}

// Team Performance Chart
function renderTeamPerformanceChart() {
    const ctx = document.getElementById('teamPerformanceChart').getContext('2d');
    
    // Calculate team performance
    const teamData = {};
    
    allUsers.forEach(user => {
        if (!teamData[user.team]) {
            teamData[user.team] = { total: 0, completed: 0, users: 0 };
        }
        
        const userSkills = allSkillData[user.username];
        if (userSkills) {
            teamData[user.team].users++;
            userSkills.forEach(skill => {
                availableMarkets.forEach(market => {
                    if (skill[market] !== undefined && skill[market] !== '') {
                        teamData[user.team].total++;
                        if (skill[market] === 'yes') {
                            teamData[user.team].completed++;
                        }
                    }
                });
            });
        }
    });
    
    const teams = Object.keys(teamData);
    const completionRates = teams.map(team => 
        teamData[team].total > 0 ? Math.round((teamData[team].completed / teamData[team].total) * 100) : 0
    );
    
    if (chartInstances.teamPerformance) {
        chartInstances.teamPerformance.destroy();
    }
    
    chartInstances.teamPerformance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: teams,
            datasets: [{
                label: 'Completion Rate (%)',
                data: completionRates,
                backgroundColor: 'rgba(48, 86, 211, 0.8)',
                borderColor: 'rgba(48, 86, 211, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100
                }
            }
        }
    });
}

// Category Progress Chart
function renderCategoryProgressChart() {
    const ctx = document.getElementById('categoryProgressChart').getContext('2d');
    
    // Calculate category progress
    const categoryData = {};
    
    categories.forEach(category => {
        categoryData[category] = { total: 0, completed: 0 };
    });
    
    Object.values(allSkillData).forEach(userSkills => {
        userSkills.forEach(skill => {
            if (categoryData[skill.Category]) {
                availableMarkets.forEach(market => {
                    if (skill[market] !== undefined && skill[market] !== '') {
                        categoryData[skill.Category].total++;
                        if (skill[market] === 'yes') {
                            categoryData[skill.Category].completed++;
                        }
                    }
                });
            }
        });
    });
    
    const categoryLabels = Object.keys(categoryData);
    const categoryProgress = categoryLabels.map(cat => 
        categoryData[cat].total > 0 ? Math.round((categoryData[cat].completed / categoryData[cat].total) * 100) : 0
    );
    
    if (chartInstances.categoryProgress) {
        chartInstances.categoryProgress.destroy();
    }
    
    chartInstances.categoryProgress = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: categoryLabels,
            datasets: [{
                data: categoryProgress,
                backgroundColor: [
                    'rgba(48, 86, 211, 0.8)',
                    'rgba(19, 194, 150, 0.8)',
                    'rgba(251, 189, 35, 0.8)',
                    'rgba(248, 114, 114, 0.8)',
                    'rgba(58, 191, 248, 0.8)'
                ],
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
}

// Populate tables
function populateTables() {
    populateTopPerformers();
    populateSkillGaps();
    populateTeamComparison();
}

// Top Performers
function populateTopPerformers() {
    const userPerformance = [];
    
    allUsers.forEach(user => {
        const userSkills = allSkillData[user.username];
        if (userSkills) {
            let total = 0;
            let completed = 0;
            
            userSkills.forEach(skill => {
                availableMarkets.forEach(market => {
                    if (skill[market] !== undefined && skill[market] !== '') {
                        total++;
                        if (skill[market] === 'yes') {
                            completed++;
                        }
                    }
                });
            });
            
            if (total > 0) {
                userPerformance.push({
                    name: user.name,
                    team: user.team,
                    total: total,
                    completed: completed,
                    percentage: Math.round((completed / total) * 100)
                });
            }
        }
    });
    
    // Sort by percentage
    userPerformance.sort((a, b) => b.percentage - a.percentage);
    
    // Display top 10
    const tbody = document.getElementById('topPerformersBody');
    tbody.innerHTML = '';
    
    userPerformance.slice(0, 10).forEach((user, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><span class="rank-badge">${index + 1}</span></td>
            <td>${user.name}</td>
            <td>${user.team}</td>
            <td>
                <div class="progress-cell">
                    <span class="progress-text">${user.percentage}%</span>
                    <div class="progress-bar-mini">
                        <div class="progress-fill" style="width: ${user.percentage}%"></div>
                    </div>
                </div>
            </td>
            <td>${user.completed}/${user.total}</td>
        `;
        tbody.appendChild(row);
    });
}

// Skill Gaps
function populateSkillGaps() {
    const skillCoverage = {};
    
    // Calculate coverage for each skill
    taskMapping.forEach(task => {
        const key = `${task.Category}-${task.Task_Name}`;
        skillCoverage[key] = {
            category: task.Category,
            taskName: task.Task_Name,
            required: 0,
            covered: 0
        };
        
        // Count required markets
        availableMarkets.forEach(market => {
            if (task[market] === 'x') {
                skillCoverage[key].required++;
                
                // Check coverage across all users
                let marketCovered = false;
                Object.values(allSkillData).forEach(userSkills => {
                    const userTask = userSkills.find(s => s.Item_ID === task.Item_ID);
                    if (userTask && userTask[market] === 'yes') {
                        marketCovered = true;
                    }
                });
                
                if (marketCovered) {
                    skillCoverage[key].covered++;
                }
            }
        });
    });
    
    // Calculate gaps and sort
    const gaps = Object.values(skillCoverage)
        .map(skill => ({
            ...skill,
            coverage: skill.required > 0 ? Math.round((skill.covered / skill.required) * 100) : 100,
            gap: skill.required - skill.covered
        }))
        .filter(skill => skill.gap > 0)
        .sort((a, b) => b.gap - a.gap);
    
    // Display top gaps
    const tbody = document.getElementById('skillGapsBody');
    tbody.innerHTML = '';
    
    gaps.slice(0, 10).forEach(gap => {
        const priority = gap.coverage < 25 ? 'Critical' : gap.coverage < 50 ? 'High' : 'Medium';
        const priorityClass = gap.coverage < 25 ? 'danger' : gap.coverage < 50 ? 'warning' : 'info';
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${gap.taskName}</td>
            <td>${gap.category}</td>
            <td>${gap.coverage}%</td>
            <td>${gap.gap} markets</td>
            <td><span class="priority-badge ${priorityClass}">${priority}</span></td>
        `;
        tbody.appendChild(row);
    });
}

// Team Comparison
function populateTeamComparison() {
    const teamStats = {};
    
    // Calculate detailed team statistics
    allUsers.forEach(user => {
        if (!teamStats[user.team]) {
            teamStats[user.team] = {
                members: 0,
                categoryPerformance: {},
                totalSkills: 0,
                completedSkills: 0
            };
            
            categories.forEach(cat => {
                teamStats[user.team].categoryPerformance[cat] = { total: 0, completed: 0 };
            });
        }
        
        teamStats[user.team].members++;
        
        const userSkills = allSkillData[user.username];
        if (userSkills) {
            userSkills.forEach(skill => {
                availableMarkets.forEach(market => {
                    if (skill[market] !== undefined && skill[market] !== '') {
                        teamStats[user.team].totalSkills++;
                        teamStats[user.team].categoryPerformance[skill.Category].total++;
                        
                        if (skill[market] === 'yes') {
                            teamStats[user.team].completedSkills++;
                            teamStats[user.team].categoryPerformance[skill.Category].completed++;
                        }
                    }
                });
            });
        }
    });
    
    // Render comparison table
    const tbody = document.getElementById('teamComparisonBody');
    tbody.innerHTML = '';
    
    Object.entries(teamStats).forEach(([team, stats]) => {
        const avgCompletion = stats.totalSkills > 0 
            ? Math.round((stats.completedSkills / stats.totalSkills) * 100) : 0;
        
        // Find strongest and weakest categories
        let strongest = { category: '-', percentage: 0 };
        let weakest = { category: '-', percentage: 100 };
        
        Object.entries(stats.categoryPerformance).forEach(([cat, perf]) => {
            if (perf.total > 0) {
                const percentage = Math.round((perf.completed / perf.total) * 100);
                if (percentage > strongest.percentage) {
                    strongest = { category: cat, percentage };
                }
                if (percentage < weakest.percentage) {
                    weakest = { category: cat, percentage };
                }
            }
        });
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${team}</td>
            <td>${stats.members}</td>
            <td>
                <div class="progress-cell">
                    <span class="progress-text">${avgCompletion}%</span>
                    <div class="progress-bar-mini">
                        <div class="progress-fill" style="width: ${avgCompletion}%"></div>
                    </div>
                </div>
            </td>
            <td>${strongest.category} (${strongest.percentage}%)</td>
            <td>${weakest.category} (${weakest.percentage}%)</td>
            <td>
                <span class="trend ${avgCompletion >= 50 ? 'positive' : 'negative'}">
                    ${avgCompletion >= 50 ? '↑' : '↓'}
                </span>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// Market Readiness
function renderMarketReadiness() {
    const container = document.getElementById('marketCardsGrid');
    container.innerHTML = '';
    
    // Also populate market filter
    const marketFilter = document.getElementById('marketFilter');
    marketFilter.innerHTML = '<option value="all">All Markets</option>';
    
    availableMarkets.forEach(market => {
        // Add to filter
        const option = document.createElement('option');
        option.value = market;
        option.textContent = market;
        marketFilter.appendChild(option);
        
        // Calculate market readiness
        let requiredSkills = 0;
        let coveredSkills = 0;
        
        taskMapping.forEach(task => {
            if (task[market] === 'x') {
                requiredSkills++;
                
                // Check if any user has this skill for this market
                let skillCovered = false;
                Object.values(allSkillData).forEach(userSkills => {
                    const userTask = userSkills.find(s => s.Item_ID === task.Item_ID);
                    if (userTask && userTask[market] === 'yes') {
                        skillCovered = true;
                    }
                });
                
                if (skillCovered) coveredSkills++;
            }
        });
        
        const readiness = requiredSkills > 0 
            ? Math.round((coveredSkills / requiredSkills) * 100) : 0;
        
        const statusClass = readiness >= 75 ? 'success' : readiness >= 50 ? 'warning' : 'danger';
        const statusText = readiness >= 75 ? 'Ready' : readiness >= 50 ? 'Partial' : 'Not Ready';
        
        const card = document.createElement('div');
        card.className = 'market-card';
        card.innerHTML = `
            <div class="market-card-header">
                <h4>${market}</h4>
                <span class="status-badge ${statusClass}">${statusText}</span>
            </div>
            <div class="market-card-body">
                <div class="readiness-score">
                    <span class="score-value">${readiness}%</span>
                    <span class="score-label">Readiness</span>
                </div>
                <div class="market-stats">
                    <div class="stat-item">
                        <span class="stat-value">${coveredSkills}</span>
                        <span class="stat-label">Covered</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value">${requiredSkills - coveredSkills}</span>
                        <span class="stat-label">Gaps</span>
                    </div>
                </div>
            </div>
        `;
        container.appendChild(card);
    });
}

// Export functions
function exportFullReport() {
    document.getElementById('exportModal').classList.add('active');
}

function closeExportModal() {
    document.getElementById('exportModal').classList.remove('active');
}

function executeExport() {
    const wb = XLSX.utils.book_new();
    
    // Executive Summary
    if (document.getElementById('exportSummary').checked) {
        const summaryData = [
            ['Executive Summary'],
            [''],
            ['Metric', 'Value'],
            ['Total Users', Object.keys(allSkillData).length],
            ['Overall Completion', document.getElementById('overallCompletion').textContent],
            ['Total Tasks', taskMapping.length],
            ['Active Markets', availableMarkets.length]
        ];
        const ws = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(wb, ws, 'Summary');
    }
    
    // Team Performance
    if (document.getElementById('exportTeams').checked) {
        const teamData = [];
        const tbody = document.getElementById('teamComparisonBody');
        const rows = tbody.getElementsByTagName('tr');
        
        teamData.push(['Team', 'Members', 'Avg Completion', 'Strongest Category', 'Weakest Category']);
        for (let row of rows) {
            const cells = row.getElementsByTagName('td');
            teamData.push([
                cells[0].textContent,
                cells[1].textContent,
                cells[2].querySelector('.progress-text').textContent,
                cells[3].textContent,
                cells[4].textContent
            ]);
        }
        
        const ws = XLSX.utils.aoa_to_sheet(teamData);
        XLSX.utils.book_append_sheet(wb, ws, 'Teams');
    }
    
    // Save file
    XLSX.writeFile(wb, `skill_report_${new Date().getTime()}.xlsx`);
    closeExportModal();
    showToast('Report exported successfully', 'success');
}

// Utility functions
function showLoading(show, message = 'Loading...') {
    const overlay = document.getElementById('reportingLoadingOverlay');
    if (overlay) {
        overlay.classList.toggle('active', show);
        if (message) showLoadingMessage(message);
    }
}

function showLoadingMessage(message) {
    const element = document.getElementById('reportingLoadingMessage');
    if (element) element.textContent = message;
}

function showNoAccess() {
    document.getElementById('reportingPage').style.display = 'none';
    document.getElementById('reportingNoAccessMessage').style.display = 'flex';
}

function showToast(message, type = 'info') {
    const container = document.getElementById('reportingToastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    const icons = {
        success: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>',
        error: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>',
        warning: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>',
        info: '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width: 20px; height: 20px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>'
    };
    
    toast.innerHTML = `${icons[type]}<span style="margin-left: 8px;">${message}</span>`;
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.classList.add('removing');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// Additional functions
function refreshReportData() {
    initializeReporting();
    showToast('Report data refreshed', 'success');
}

function updateReportingPeriod() {
    // Placeholder for period filtering
    initializeReporting();
}

function filterMarketData() {
    const filter = document.getElementById('marketFilter').value;
    // Filter market cards based on selection
    renderMarketReadiness();
}

function toggleTeamView() {
    // Toggle between different team views
    populateTeamComparison();
}

function exportTeamData() {
    const wb = XLSX.utils.book_new();
    // Export team chart data
    XLSX.writeFile(wb, `team_performance_${new Date().getTime()}.xlsx`);
}

function exportCategoryData() {
    const wb = XLSX.utils.book_new();
    // Export category data
    XLSX.writeFile(wb, `category_progress_${new Date().getTime()}.xlsx`);
}

function exportMarketData() {
    const wb = XLSX.utils.book_new();
    // Export market readiness data
    XLSX.writeFile(wb, `market_readiness_${new Date().getTime()}.xlsx`);
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeReporting);
} else {
    initializeReporting();
}