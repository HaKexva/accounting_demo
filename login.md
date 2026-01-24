---
layout: default
title: Login
---

<div class="login-container">
  <h1 style="text-align: center; margin-bottom: 30px;">登入系統</h1>
  
  <div id="error-message" class="error-message"></div>
  
  <div id="buttonDiv"></div>
  
  <p class="security-note">
    🔒 您的資料安全受到保護<br>
    只有授權用戶可以訪問此系統
  </p>
</div>


<script>
  // Check if user is already logged in
  const existingSession = localStorage.getItem('auth_session');
  if (existingSession) {
    try {
      const session = JSON.parse(existingSession);
      if (session.expiresAt && session.expiresAt > Date.now()) {
        // User is already logged in, redirect to home
        window.location.href = '/accounting/';
      }
    } catch (e) {
      localStorage.removeItem('auth_session');
    }
  }

  // Load auth script
  const script = document.createElement('script');
  script.src = '/accounting/assets/auth.js';
  script.onload = function() {
    // Initialize auth and render button
    authManager.init().then(() => {
      authManager.renderSignInButton('buttonDiv');
    });
  };
  document.head.appendChild(script);
  
  // Check for error in URL params
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('error') === 'unauthorized') {
    const errorEl = document.getElementById('error-message');
    errorEl.textContent = '您的帳號未被授權訪問此系統';
    errorEl.style.display = 'block';
  }
</script>