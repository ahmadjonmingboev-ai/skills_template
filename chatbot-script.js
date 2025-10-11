// Chatbot Widget JavaScript
(function() {
    'use strict';

    // Configuration
    const config = {
        workspaceUrl: 'YOUR_WORKSPACE_AI_URL_HERE', // Replace with your actual workspace AI URL
        workspaceName: 'Workspace AI Assistant',
        loadDelay: 500,
        animationDuration: 300
    };

    // State management
    const state = {
        isOpen: false,
        isMaximized: false,
        isIframeLoaded: false,
        iframeUrl: null
    };

    // DOM Elements
    let elements = {};

    // Initialize the chatbot widget
    function init() {
        // Wait for DOM to be ready
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initialize);
        } else {
            initialize();
        }
    }

    function initialize() {
        cacheElements();
        setupEventListeners();
        updateWorkspaceName();
        restoreState();
    }

    // Cache DOM elements for better performance
    function cacheElements() {
        elements = {
            trigger: document.getElementById('chatbot-trigger'),
            modal: document.getElementById('chatbot-modal'),
            iframe: document.getElementById('chatbot-iframe'),
            loader: document.querySelector('.iframe-loader'),
            closeBtn: document.getElementById('btn-close'),
            maximizeBtn: document.getElementById('btn-maximize'),
            externalBtn: document.getElementById('btn-external'),
            iconChat: document.querySelector('.icon-chat'),
            iconClose: document.querySelector('.icon-close'),
            iconMaximize: document.querySelector('.icon-maximize'),
            iconMinimize: document.querySelector('.icon-minimize'),
            headerTitle: document.querySelector('.header-title'),
            chatLabel: document.querySelector('.chatbot-label')
        };
    }

    // Setup all event listeners
    function setupEventListeners() {
        // Trigger button click
        elements.trigger.addEventListener('click', toggleChatbot);

        // Header buttons
        elements.closeBtn.addEventListener('click', minimizeChatbot);
        elements.maximizeBtn.addEventListener('click', toggleMaximize);
        elements.externalBtn.addEventListener('click', openExternal);

        // Iframe load event
        elements.iframe.addEventListener('load', handleIframeLoad);

        // Keyboard shortcuts
        document.addEventListener('keydown', handleKeyboard);

        // Click outside to close (optional)
        document.addEventListener('click', handleOutsideClick);

        // Handle window resize
        window.addEventListener('resize', handleResize);

        // Prevent iframe from being refreshed on modal close
        elements.modal.addEventListener('transitionend', handleTransitionEnd);
    }

    // Update workspace name from config
    function updateWorkspaceName() {
        if (elements.headerTitle) {
            elements.headerTitle.textContent = config.workspaceName;
        }
    }

    // Toggle chatbot open/close
    function toggleChatbot(e) {
        e.stopPropagation();
        
        if (state.isOpen) {
            minimizeChatbot();
        } else {
            openChatbot();
        }
    }

    // Open chatbot modal
    function openChatbot() {
        if (state.isOpen) return;

        state.isOpen = true;
        
        // Update trigger button
        elements.trigger.classList.add('active');
        elements.iconChat.style.display = 'none';
        elements.iconClose.style.display = 'block';
        
        // Show modal with animation
        elements.modal.classList.add('active', 'animating');
        
        // Load iframe only once
        if (!state.isIframeLoaded && config.workspaceUrl) {
            loadIframe();
        } else if (state.iframeUrl) {
            // Iframe already loaded, just show loader briefly
            showLoader();
            setTimeout(hideLoader, 300);
        }

        // Save state
        saveState();
        
        // Announce to screen readers
        announceState('Chatbot opened');
    }

    // Minimize chatbot (not close/refresh)
    function minimizeChatbot() {
        if (!state.isOpen) return;

        state.isOpen = false;
        
        // Update trigger button
        elements.trigger.classList.remove('active');
        elements.iconChat.style.display = 'block';
        elements.iconClose.style.display = 'none';
        
        // Hide modal
        elements.modal.classList.remove('active');
        
        // Reset maximized state
        if (state.isMaximized) {
            state.isMaximized = false;
            elements.modal.classList.remove('maximized');
            updateMaximizeIcon();
        }

        // Save state
        saveState();
        
        // Announce to screen readers
        announceState('Chatbot minimized');
    }

    // Toggle maximize/restore
    function toggleMaximize() {
        state.isMaximized = !state.isMaximized;
        
        if (state.isMaximized) {
            elements.modal.classList.add('maximized');
            announceState('Chatbot maximized');
        } else {
            elements.modal.classList.remove('maximized');
            announceState('Chatbot restored');
        }
        
        updateMaximizeIcon();
        saveState();
    }

    // Update maximize button icon
    function updateMaximizeIcon() {
        if (state.isMaximized) {
            elements.iconMaximize.style.display = 'none';
            elements.iconMinimize.style.display = 'block';
        } else {
            elements.iconMaximize.style.display = 'block';
            elements.iconMinimize.style.display = 'none';
        }
    }

    // Open in external window
    function openExternal() {
        if (!config.workspaceUrl) {
            alert('Workspace URL not configured. Please set the workspaceUrl in the configuration.');
            return;
        }
        
        // Open in new window with specific dimensions
        const width = 1200;
        const height = 800;
        const left = (screen.width - width) / 2;
        const top = (screen.height - height) / 2;
        
        window.open(
            config.workspaceUrl,
            'WorkspaceAI',
            `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`
        );
        
        // Optionally minimize the modal
        // minimizeChatbot();
    }

    // Load iframe content
    function loadIframe() {
        if (!config.workspaceUrl) {
            console.error('Workspace URL not configured');
            hideLoader();
            showError('Workspace URL not configured. Please contact your administrator.');
            return;
        }

        showLoader();
        state.iframeUrl = config.workspaceUrl;
        elements.iframe.src = state.iframeUrl;
        state.isIframeLoaded = true;
    }

    // Handle iframe load completion
    function handleIframeLoad() {
        if (state.isIframeLoaded) {
            setTimeout(hideLoader, config.loadDelay);
        }
    }

    // Show/hide loader
    function showLoader() {
        elements.loader.classList.remove('hidden');
    }

    function hideLoader() {
        elements.loader.classList.add('hidden');
    }

    // Show error message
    function showError(message) {
        const errorDiv = document.createElement('div');
        errorDiv.className = 'iframe-error';
        errorDiv.style.cssText = 'position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); text-align: center; padding: 20px; color: #EF4444;';
        errorDiv.innerHTML = `
            <svg width="48" height="48" style="margin: 0 auto 16px;" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
            </svg>
            <p style="font-size: 14px; margin: 0;">${message}</p>
        `;
        
        elements.modal.querySelector('.chatbot-body').appendChild(errorDiv);
    }

    // Handle keyboard shortcuts
    function handleKeyboard(e) {
        // Escape key to close
        if (e.key === 'Escape' && state.isOpen) {
            minimizeChatbot();
        }
        
        // Ctrl/Cmd + Shift + C to toggle chatbot
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'C') {
            e.preventDefault();
            toggleChatbot(e);
        }
    }

    // Handle clicks outside modal
    function handleOutsideClick(e) {
        // Optional: Close modal when clicking outside
        // Uncomment if you want this behavior
        /*
        if (state.isOpen && !elements.modal.contains(e.target) && !elements.trigger.contains(e.target)) {
            minimizeChatbot();
        }
        */
    }

    // Handle window resize
    function handleResize() {
        if (state.isOpen && window.innerWidth < 768) {
            // Auto-adjust for mobile
            if (!elements.modal.classList.contains('maximized')) {
                // Optional: Auto-maximize on small screens
                // toggleMaximize();
            }
        }
    }

    // Handle transition end
    function handleTransitionEnd(e) {
        if (e.target === elements.modal) {
            elements.modal.classList.remove('animating');
        }
    }

    // Save state to sessionStorage
    function saveState() {
        try {
            sessionStorage.setItem('chatbotState', JSON.stringify({
                isOpen: state.isOpen,
                isMaximized: state.isMaximized
            }));
        } catch (e) {
            console.warn('Could not save chatbot state:', e);
        }
    }

    // Restore state from sessionStorage
    function restoreState() {
        try {
            const savedState = sessionStorage.getItem('chatbotState');
            if (savedState) {
                const parsed = JSON.parse(savedState);
                
                // Restore state after a brief delay
                setTimeout(() => {
                    if (parsed.isOpen) {
                        openChatbot();
                        
                        if (parsed.isMaximized) {
                            setTimeout(() => {
                                toggleMaximize();
                            }, 300);
                        }
                    }
                }, 100);
            }
        } catch (e) {
            console.warn('Could not restore chatbot state:', e);
        }
    }

    // Announce state changes for accessibility
    function announceState(message) {
        const announcement = document.createElement('div');
        announcement.setAttribute('role', 'status');
        announcement.setAttribute('aria-live', 'polite');
        announcement.style.position = 'absolute';
        announcement.style.left = '-9999px';
        announcement.textContent = message;
        
        document.body.appendChild(announcement);
        
        setTimeout(() => {
            document.body.removeChild(announcement);
        }, 1000);
    }

    // Public API (optional - for external control)
    window.WorkspaceAIChatbot = {
        open: openChatbot,
        close: minimizeChatbot,
        toggle: toggleChatbot,
        maximize: () => {
            if (!state.isMaximized) toggleMaximize();
        },
        minimize: () => {
            if (state.isMaximized) toggleMaximize();
        },
        setUrl: (url) => {
            config.workspaceUrl = url;
            if (state.isIframeLoaded) {
                state.isIframeLoaded = false;
                state.iframeUrl = null;
                elements.iframe.src = '';
                if (state.isOpen) {
                    loadIframe();
                }
            }
        },
        getState: () => ({...state})
    };

    // Initialize the widget
    init();

})();

// Configuration helper for Confluence
// Add this to your Confluence HTML macro
/*
<script>
document.addEventListener('DOMContentLoaded', function() {
    // Configure your workspace URL here
    if (window.WorkspaceAIChatbot) {
        window.WorkspaceAIChatbot.setUrl('https://your-workspace-ai-url.com');
    }
});
</script>
*/
