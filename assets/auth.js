// Google Authentication Module
const AUTH_CONFIG = {
  // List of allowed email addresses - modify this to control access
  ALLOWED_EMAILS: [
    'ray120424@gmail.com',
    'tubaxenor@gmail.com',
    'eva811109@gmail.com',
    // Add more authorized emails here by adding them to this array
    // Example: 'another-user@gmail.com',
  ],
  
  // Pages that don't require authentication
  PUBLIC_PAGES: [
    '/accounting/login.html',
    '/accounting/unauthorized.html'
  ],
  
  // Google OAuth Client ID
  CLIENT_ID: '196313952790-v65pffkno3ulbbu4337geb36939bv9di.apps.googleusercontent.com',
  
  // Session duration in milliseconds (24 hours)
  SESSION_DURATION: 24 * 60 * 60 * 1000
};

// Authentication state management
class AuthManager {
  constructor() {
    this.currentUser = null;
    this.isInitialized = false;
  }

  // Initialize Google Sign-In
  async init() {
    if (this.isInitialized) return;
    if (this.isLocalhost()) {
      this.isInitialized = true;
      return;
    }
    
    return new Promise((resolve) => {
      if (typeof google !== 'undefined' && google.accounts) {
        google.accounts.id.initialize({
          client_id: AUTH_CONFIG.CLIENT_ID,
          callback: this.handleCredentialResponse.bind(this),
          auto_select: true,
          cancel_on_tap_outside: false
        });
        this.isInitialized = true;
        resolve();
      } else {
        // Wait for Google Sign-In library to load
        setTimeout(() => this.init().then(resolve), 100);
      }
    });
  }

  // Check if the current page is localhost
  isLocalhost() {
    return window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
  }

  // Handle Google Sign-In response
  handleCredentialResponse(response) {
    // Decode JWT token to get user information
    const userInfo = this.parseJwt(response.credential);
    
    // Check if user is authorized
    if (this.isEmailAllowed(userInfo.email)) {
      // Store user session
      this.setUserSession({
        email: userInfo.email,
        name: userInfo.name,
        picture: userInfo.picture,
        token: response.credential,
        expiresAt: Date.now() + AUTH_CONFIG.SESSION_DURATION
      });
      
      // Redirect to original page or home
      const redirectUrl = sessionStorage.getItem('auth_redirect') || '/accounting/';
      sessionStorage.removeItem('auth_redirect');
      window.location.href = redirectUrl;
    } else {
      // User not authorized
      this.clearSession();
      window.location.href = '/accounting/unauthorized.html';
    }
  }

  // Parse JWT token
  parseJwt(token) {
    const base64Url = token.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
      return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));
    return JSON.parse(jsonPayload);
  }

  // Check if email is in allowed list
  isEmailAllowed(email) {
    return AUTH_CONFIG.ALLOWED_EMAILS.includes(email.toLowerCase());
  }

  // Set user session in localStorage
  setUserSession(userData) {
    localStorage.setItem('auth_session', JSON.stringify(userData));
    this.currentUser = userData;
  }

  // Get current user session
  getUserSession() {
    if (this.currentUser) return this.currentUser;
    
    const sessionStr = localStorage.getItem('auth_session');
    if (!sessionStr) return null;
    
    try {
      const session = JSON.parse(sessionStr);
      
      // Check if session is expired
      if (session.expiresAt && session.expiresAt < Date.now()) {
        this.clearSession();
        return null;
      }
      
      this.currentUser = session;
      return session;
    } catch (e) {
      this.clearSession();
      return null;
    }
  }

  // Clear user session
  clearSession() {
    localStorage.removeItem('auth_session');
    this.currentUser = null;
  }

  // Sign out user
  signOut() {
    this.clearSession();
    if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
      google.accounts.id.disableAutoSelect();
    }
    window.location.href = '/accounting/login.html';
  }

  // Check if current page requires authentication
  isProtectedPage() {
    if (this.isLocalhost()) {
      return false;
    }
    const currentPath = window.location.pathname;
    return !AUTH_CONFIG.PUBLIC_PAGES.some(page => 
      currentPath === page || currentPath.endsWith(page)
    );
  }

  // Render sign-in button
  renderSignInButton(elementId = 'buttonDiv') {
    if (this.isLocalhost()) {
      return;
    }
    const element = document.getElementById(elementId);
    if (element && typeof google !== 'undefined' && google.accounts) {
      google.accounts.id.renderButton(element, {
        theme: 'outline',
        size: 'large',
        text: 'signin_with',
        width: 250
      });
    }
  }

  // Show One Tap prompt
  showOneTap() {
    if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
      google.accounts.id.prompt((notification) => {
        if (notification.isNotDisplayed() || notification.isSkippedMoment()) {
          // One Tap not displayed or skipped, show the button
          this.renderSignInButton();
        }
      });
    }
  }
}

// Create global auth instance
const authManager = new AuthManager();

// Protection logic for all pages
async function protectPage() {
  // Wait for auth manager to initialize
  await authManager.init();
  
  // Check if page requires authentication
  if (authManager.isProtectedPage()) {
    const session = authManager.getUserSession();
    
    if (!session) {
      // Store current URL for redirect after login
      sessionStorage.setItem('auth_redirect', window.location.pathname);
      // Redirect to login page
      window.location.href = '/accounting/login.html';
      return false;
    }
    
    // User is authenticated, add logout button to hamburger menu (same for all devices)
    const siteNav = document.querySelector('.site-nav .trigger') || document.querySelector('.site-nav');
    if (siteNav) {
      // Check if logout button has already been added
      let logoutLink = siteNav.querySelector('.logout-link');
      if (!logoutLink) {
        logoutLink = document.createElement('a');
        logoutLink.className = 'page-link logout-link';
        logoutLink.href = '#';
        logoutLink.textContent = '登出';
        logoutLink.onclick = (e) => {
          e.preventDefault();
          authManager.signOut();
        };
        
        siteNav.appendChild(logoutLink);
      }
    }
    
    // Hide original user-info (no longer displayed in top right)
    const userInfoEl = document.getElementById('user-info');
    if (userInfoEl) {
      userInfoEl.style.display = 'none';
    }
    
    return true;
  }
  
  return true;
}

// Make Accounting logo link to homepage
function setupHomeLink() {
  const siteTitle = document.querySelector('.site-title');
  if (siteTitle) {
    siteTitle.style.cursor = 'pointer';
    siteTitle.href = '/accounting/';
  }
}

// Auto-protect pages on load
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    protectPage();
    setupHomeLink();
  });
} else {
  protectPage();
  setupHomeLink();
}

// Export for use in other scripts
window.authManager = authManager;
window.AUTH_CONFIG = AUTH_CONFIG;
window.protectPage = protectPage; // Export protectPage function for use by other scripts