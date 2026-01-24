// Sync Status Module - Unified caching and sync status indicator
// This module provides IndexedDB caching and a nav sync status icon

const SyncStatus = (function() {
  // ===== IndexedDB Configuration =====
  const DB_NAME = 'AccountingAppCache';
  const DB_VERSION = 2;
  const STORE_NAME = 'cache';

  let dbInstance = null;

  // Sync state
  let syncState = {
    status: 'idle', // 'idle', 'syncing', 'synced', 'error'
    lastSyncTime: null,
    pendingOperations: 0
  };

  // ===== IndexedDB Operations =====
  const openDB = () => {
    return new Promise((resolve, reject) => {
      if (dbInstance) {
        resolve(dbInstance);
        return;
      }

      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onerror = () => reject(request.error);

      request.onsuccess = () => {
        dbInstance = request.result;
        resolve(dbInstance);
      };

      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME, { keyPath: 'key' });
        }
      };
    });
  };

  const getFromCache = async (key) => {
    try {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const request = store.get(key);
        request.onsuccess = () => resolve(request.result?.value || null);
        request.onerror = () => reject(request.error);
      });
    } catch (e) {
      console.warn('[SyncStatus] getFromCache error:', e);
      return null;
    }
  };

  const setToCache = async (key, value) => {
    try {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        const request = store.put({ key, value, timestamp: Date.now() });
        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
      });
    } catch (e) {
      console.warn('[SyncStatus] setToCache error:', e);
      return false;
    }
  };

  const getCacheTimestamp = async (key) => {
    try {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const request = store.get(key);
        request.onsuccess = () => resolve(request.result?.timestamp || null);
        request.onerror = () => reject(request.error);
      });
    } catch (e) {
      return null;
    }
  };

  const clearCache = async () => {
    try {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        const request = store.clear();
        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
      });
    } catch (e) {
      console.warn('[SyncStatus] clearCache error:', e);
      return false;
    }
  };

  // ===== Sync Status UI =====
  let statusIcon = null;

  // Font Awesome icon classes for each state
  const ICONS = {
    idle: 'fa-solid fa-cloud',
    syncing: 'fa-solid fa-arrows-rotate',
    synced: 'fa-solid fa-check',
    error: 'fa-solid fa-download'
  };

  const createStatusIcon = () => {
    if (statusIcon) return statusIcon;

    // Create the sync status icon container
    statusIcon = document.createElement('div');
    statusIcon.id = 'sync-status-icon';
    statusIcon.className = 'sync-status-icon';
    statusIcon.innerHTML = `
      <i class="sync-icon ${ICONS.idle}"></i>
      <span class="sync-tooltip"></span>
    `;

    // Add styles
    const style = document.createElement('style');
    style.textContent = `
      .sync-status-icon {
        position: relative;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 32px;
        height: 32px;
        cursor: pointer;
        margin-right: 8px;
        border-radius: 50%;
        transition: background-color 0.2s;
      }

      .sync-status-icon:hover {
        background-color: rgba(0, 0, 0, 0.05);
      }

      .sync-status-icon .sync-icon {
        font-size: 16px;
        color: #9e9e9e;
        transition: color 0.3s;
      }

      .sync-status-icon.syncing .sync-icon {
        color: #2196F3;
        animation: sync-spin 1s linear infinite;
      }

      .sync-status-icon.synced .sync-icon {
        color: #4CAF50;
      }

      .sync-status-icon.error .sync-icon {
        color: #2196F3;
      }

      .sync-status-icon.idle .sync-icon {
        color: #9e9e9e;
      }

      @keyframes sync-spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }

      .sync-status-icon .sync-tooltip {
        position: absolute;
        top: 100%;
        right: 0;
        margin-top: 8px;
        padding: 6px 12px;
        background: rgba(0, 0, 0, 0.8);
        color: #fff;
        font-size: 12px;
        border-radius: 4px;
        white-space: nowrap;
        opacity: 0;
        visibility: hidden;
        transition: opacity 0.2s, visibility 0.2s;
        z-index: 1000;
        pointer-events: none;
      }

      .sync-status-icon:hover .sync-tooltip {
        opacity: 1;
        visibility: visible;
      }

      .sync-status-icon .sync-tooltip::before {
        content: '';
        position: absolute;
        bottom: 100%;
        right: 10px;
        border: 6px solid transparent;
        border-bottom-color: rgba(0, 0, 0, 0.8);
      }

      /* Mobile adjustments */
      @media (max-width: 600px) {
        .sync-status-icon {
          width: 28px;
          height: 28px;
          margin-right: 4px;
        }

        .sync-status-icon .sync-icon {
          font-size: 14px;
        }
      }
    `;

    if (!document.getElementById('sync-status-styles')) {
      style.id = 'sync-status-styles';
      document.head.appendChild(style);
    }

    return statusIcon;
  };

  const formatTimestamp = (timestamp) => {
    if (!timestamp) return '從未同步'; // Keep original Chinese text for UI display
    const date = new Date(timestamp);
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);

    if (diffMins < 1) return '剛剛同步';
    if (diffMins < 60) return `${diffMins} 分鐘前同步`;
    if (diffHours < 24) return `${diffHours} 小時前同步`;

    return date.toLocaleString('zh-TW', {
      month: 'numeric',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }) + ' 同步';
  };

  const updateStatusDisplay = () => {
    if (!statusIcon) return;

    // Remove all status classes
    statusIcon.classList.remove('idle', 'syncing', 'synced', 'error');

    // Add current status class
    statusIcon.classList.add(syncState.status);

    // Update icon class
    const iconEl = statusIcon.querySelector('.sync-icon');
    if (iconEl) {
      // Remove all icon classes and add the correct one
      iconEl.className = `sync-icon ${ICONS[syncState.status] || ICONS.idle}`;
    }

    // Update tooltip
    const tooltip = statusIcon.querySelector('.sync-tooltip');
    if (tooltip) {
      let message = '';
      switch (syncState.status) {
        case 'syncing':
          message = '正在同步...';
          if (syncState.pendingOperations > 0) {
            message += ` (${syncState.pendingOperations} 項)`;
          }
          break;
        case 'synced':
          message = formatTimestamp(syncState.lastSyncTime);
          break;
        case 'error':
          message = '點擊重新載入';
          break;
        default:
          message = syncState.lastSyncTime ? formatTimestamp(syncState.lastSyncTime) : '未同步';
      }
      tooltip.textContent = message;
    }
  };

  const insertStatusIcon = () => {
    const icon = createStatusIcon();

    // Try to insert into site navigation
    const siteNav = document.querySelector('.site-nav');
    if (siteNav) {
      // Insert before the nav-trigger (hamburger menu) or at the start
      const navTrigger = siteNav.querySelector('.nav-trigger');
      const menuIcon = siteNav.querySelector('label[for="nav-trigger"]');

      if (menuIcon) {
        siteNav.insertBefore(icon, menuIcon);
      } else if (navTrigger) {
        siteNav.insertBefore(icon, navTrigger);
      } else {
        siteNav.insertBefore(icon, siteNav.firstChild);
      }
    } else {
      // Fallback: insert into header
      const header = document.querySelector('.site-header');
      if (header) {
        const wrapper = header.querySelector('.wrapper');
        if (wrapper) {
          wrapper.appendChild(icon);
        }
      }
    }

    updateStatusDisplay();
  };

  // ===== Public API =====
  const startSync = (operationCount = 1) => {
    syncState.status = 'syncing';
    syncState.pendingOperations += operationCount;
    updateStatusDisplay();
  };

  const endSync = (success = true, operationCount = 1) => {
    syncState.pendingOperations = Math.max(0, syncState.pendingOperations - operationCount);

    if (syncState.pendingOperations === 0) {
      if (success) {
        syncState.status = 'synced';
        syncState.lastSyncTime = Date.now();
        // Save last sync time
        setToCache('_lastSyncTime', syncState.lastSyncTime).catch(() => {});
      } else {
        syncState.status = 'error';
      }
    }

    updateStatusDisplay();
  };

  const setError = () => {
    syncState.status = 'error';
    syncState.pendingOperations = 0;
    updateStatusDisplay();
  };

  const setIdle = () => {
    syncState.status = 'idle';
    updateStatusDisplay();
  };

  const getLastSyncTime = () => syncState.lastSyncTime;

  // Initialize when DOM is ready
  const init = async () => {
    // Load last sync time from cache
    try {
      const lastSync = await getFromCache('_lastSyncTime');
      if (lastSync) {
        syncState.lastSyncTime = lastSync;
        syncState.status = 'synced';
      }
    } catch (e) {
      // Ignore
    }

    // Sync icon removed from navigation - keep only cache functionality
  };

  // Auto-initialize
  init();

  // Return public API
  return {
    // Cache operations
    getFromCache,
    setToCache,
    getCacheTimestamp,
    clearCache,

    // Sync status
    startSync,
    endSync,
    setError,
    setIdle,
    getLastSyncTime,

    // Utilities
    formatTimestamp,

    // For debugging
    getState: () => ({ ...syncState })
  };
})();

// Export for global access
window.SyncStatus = SyncStatus;
