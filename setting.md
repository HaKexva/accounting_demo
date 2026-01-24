---
layout: page
title: 設定
permalink: /settings/
---

<link rel="stylesheet" href="{{ '/assets/settings.css' | relative_url }}?v={{ site.time | date: '%s' }}">

<div id="user-info"></div>
<div id="loading-overlay" class="loading-overlay">
  <div id="loading-progress" style="width: 300px; text-align: center;">
    <div class="progress-text" style="font-size: 16px; margin-bottom: 15px; color: #333;">載入中...</div>
    <div style="width: 100%; height: 8px; background-color: #e0e0e0; border-radius: 4px; overflow: hidden;">
      <div class="progress-bar" style="width: 0%; height: 100%; background: linear-gradient(90deg, #3498db, #2ecc71); border-radius: 4px; transition: width 0.3s ease;"></div>
    </div>
  </div>
</div>
<div id="settings-root"></div>

<script>
const baseExpense = "https://script.google.com/macros/s/AKfycbxpBh0QVSVTjylhh9cj7JG9d6aJi7L7y6pQPW88EbAsNtcd5ckucLagH8XpSAGa8IZt/exec";
const baseBudget  = "https://script.google.com/macros/s/AKfycbxkOKU5YxZWP1XTCFCF7a62Ar71fUz4Qw7tjF3MvMGkLTt6QzzhGLnDsD7wVI_cgpAR/exec";

// Settings items data (titles used to identify columns in Google Sheet)
// Button colors match homepage order: blue (payment), purple (category), orange (platform)
const settingsItems = [
  { id: 'payment', icon: '💳', title: '支付方式', colorClass: 'btn-payment', headerKeywords: ['支付方式'] },
  { id: 'category', icon: '🏷️', title: '消費類別', colorClass: 'btn-category', headerKeywords: ['支出－項目', '支出-項目', '消費類別', '類別'] },
  { id: 'platform', icon: '🏦', title: '支付平台', colorClass: 'btn-platform', headerKeywords: ['支付平台', '平台'] }
];

// Dropdown data cache
let dropdownData = {};
// Column index mapping (read from header row)
let columnMapping = {};
// Original data (for sync)
let originalSheetData = [];
// Whether dropdown data has been loaded
let dropdownLoaded = false;
// Original dropdown data (for comparing changes)
let originalDropdownData = {};
// Whether there are unsaved changes
let hasUnsavedChanges = {};

const root = document.getElementById('settings-root');
const loadingOverlay = document.getElementById('loading-overlay');

// Update progress bar
function updateProgress(current, total, text = '載入中...') {
  const progressContainer = document.getElementById('loading-progress');
  if (!progressContainer) return;

  const percentage = total > 0 ? Math.min(100, Math.round((current / total) * 100)) : 0;
  const progressBar = progressContainer.querySelector('.progress-bar');
  const progressText = progressContainer.querySelector('.progress-text');

  if (progressBar) {
    progressBar.style.width = `${percentage}%`;
  }
  if (progressText) {
    progressText.textContent = total > 0 ? `${text} (${current}/${total})` : text;
  }
}

// Show/hide loading
function showLoading() {
  loadingOverlay.style.display = 'flex';
  updateProgress(0, 0, '載入中...');
}
function hideLoading() {
  const progressContainer = document.getElementById('loading-progress');
  if (progressContainer) {
    const progressBar = progressContainer.querySelector('.progress-bar');
    const progressText = progressContainer.querySelector('.progress-text');
    if (progressBar) {
      progressBar.style.width = '100%';
      progressBar.style.background = 'linear-gradient(90deg, #2ecc71, #27ae60)';
    }
    if (progressText) {
      progressText.textContent = '✓ 已完成';
      progressText.style.color = '#2ecc71';
      progressText.style.fontWeight = 'bold';
    }
  }

  setTimeout(() => {
    loadingOverlay.style.display = 'none';
    // Reset progress bar
    if (progressContainer) {
      const progressBar = progressContainer.querySelector('.progress-bar');
      const progressText = progressContainer.querySelector('.progress-text');
      if (progressBar) {
        progressBar.style.width = '0%';
        progressBar.style.background = 'linear-gradient(90deg, #3498db, #2ecc71)';
      }
      if (progressText) {
        progressText.textContent = '載入中...';
        progressText.style.color = '#333';
        progressText.style.fontWeight = 'normal';
      }
    }
  }, 800);
}

// Load dropdown data from Google Sheet (sheetIndex = 1)
async function loadDropdownData() {
  try {
    const params = { name: "Show Tab Data", sheet: 1, _t: Date.now() };
    const url = `${baseExpense}?${new URLSearchParams(params)}`;
    const res = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store"
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const responseData = await res.json();
    // Handle different data formats
    let data = null;

    // If it's an array, use directly
    if (Array.isArray(responseData)) {
      data = responseData;
    }
    // If it's an object, might be named range format, try to find first array value
    else if (typeof responseData === 'object' && responseData !== null) {
      // Find the first key with array value
      for (const key in responseData) {
        if (Array.isArray(responseData[key]) && responseData[key].length > 0) {
          data = responseData[key];
          break;
        }
      }
      // If not found, try to use entire sheet data range
      // This requires GAS to return entire sheet data, not named range
    }

    if (!data || !Array.isArray(data) || data.length === 0) {
      // Keep dropdownLoaded as false, will retry when clicking category again
      return {};
    }

    originalSheetData = data;
    const headerRow = data[0];
    // Find column index for each settings item based on header
    settingsItems.forEach(item => {
      columnMapping[item.id] = -1;
      for (let col = 0; col < headerRow.length; col++) {
        const headerText = (headerRow[col] || '').toString().trim();
        if (item.headerKeywords.some(keyword => headerText.includes(keyword))) {
          columnMapping[item.id] = col;
          break;
        }
      }
    });

    // Read options (skip header row)
    settingsItems.forEach(item => {
      dropdownData[item.id] = [];
      const col = columnMapping[item.id];
      if (col >= 0) {
        const seen = new Set(); // Avoid duplicates
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row) continue;
          const raw = row[col];
          if (raw === undefined || raw === null) continue;
          const val = raw.toString().trim();
          if (!val || seen.has(val)) continue;
          seen.add(val);
          dropdownData[item.id].push(val);
        }
      }
    });
    dropdownLoaded = true;
    return dropdownData;
  } catch (error) {
    // Keep dropdownLoaded as false, will retry when clicking category again
    return {};
  }
}

// Create main menu
function createMainMenu() {
  const container = document.createElement('div');
  container.id = 'main-menu';
  container.className = 'settings-container';

  // Add title
  const title = document.createElement('h1');
  title.className = 'settings-title';
  title.textContent = '設定';
  container.appendChild(title);

  // Add buttons container
  const buttonsContainer = document.createElement('div');
  buttonsContainer.className = 'settings-buttons-container';

  settingsItems.forEach(item => {
    const btn = document.createElement('a');
    btn.href = '#';
    btn.className = `settings-btn ${item.colorClass}`;
    btn.innerHTML = `<span class="btn-icon">${item.icon}</span>${item.title}`;
    btn.onclick = (e) => {
      e.preventDefault();
      showPanel(item.id);
    };
    buttonsContainer.appendChild(btn);
  });

  container.appendChild(buttonsContainer);
  return container;
}

// Create content panel
function createPanel(item) {
  const panel = document.createElement('div');
  panel.id = `panel-${item.id}`;
  panel.className = 'content-panel';
  panel.dataset.panel = item.id; // For panel-specific background colors

  panel.innerHTML = `
    <div class="panel-header">
      <h2 class="panel-title">${item.icon} ${item.title}</h2>
      <button class="back-btn">← 返回</button>
    </div>
    <div class="panel-content">
      <div class="options-list" id="list-${item.id}"></div>
      <div class="add-option-container">
        <input type="text" class="add-option-input" id="input-${item.id}" name="input-${item.id}" placeholder="新增選項...">
        <button class="add-option-btn" id="add-btn-${item.id}">+ 新增</button>
      </div>
      <div class="save-container">
        <button class="save-btn" id="save-btn-${item.id}" disabled>修改</button>
        <span class="unsaved-hint" id="unsaved-hint-${item.id}"></span>
      </div>
    </div>
  `;

  panel.querySelector('.back-btn').onclick = () => handleBackClick(item.id);
  panel.querySelector(`#add-btn-${item.id}`).onclick = () => addOption(item.id);
  panel.querySelector(`#input-${item.id}`).onkeypress = (e) => {
    if (e.key === 'Enter') addOption(item.id);
  };
  panel.querySelector(`#save-btn-${item.id}`).onclick = () => saveChanges(item.id);

  return panel;
}

// Render options list
function renderOptionsList(itemId) {
  const list = document.getElementById(`list-${itemId}`);
  const options = dropdownData[itemId] || [];

  if (options.length === 0) {
    list.innerHTML = '<div class="empty-msg">尚無選項</div>';
    return;
  }

  list.innerHTML = '';
  options.forEach((option, index) => {
    const optionDiv = document.createElement('div');
    optionDiv.className = 'option-item';
    optionDiv.dataset.index = index;

    // Reorder buttons container
    const reorderBtns = document.createElement('div');
    reorderBtns.className = 'reorder-btns';

    // Move up button
    const upBtn = document.createElement('button');
    upBtn.className = 'reorder-btn';
    upBtn.textContent = '▲';
    upBtn.title = 'Move up';
    upBtn.disabled = index === 0;
    upBtn.onclick = () => moveOption(itemId, index, index - 1);

    // Order number
    const orderNum = document.createElement('span');
    orderNum.className = 'order-number';
    orderNum.textContent = index + 1;

    // Move down button
    const downBtn = document.createElement('button');
    downBtn.className = 'reorder-btn';
    downBtn.textContent = '▼';
    downBtn.title = 'Move down';
    downBtn.disabled = index === options.length - 1;
    downBtn.onclick = () => moveOption(itemId, index, index + 1);

    reorderBtns.appendChild(upBtn);
    reorderBtns.appendChild(orderNum);
    reorderBtns.appendChild(downBtn);

    // Text display
    const textSpan = document.createElement('span');
    textSpan.className = 'option-text';
    textSpan.textContent = option;

    // Edit button (pen icon)
    const editBtn = document.createElement('button');
    editBtn.className = 'edit-option-btn';
    editBtn.innerHTML = '<i class="fa-solid fa-pen"></i>';
    editBtn.title = '編輯';
    editBtn.onclick = () => startEdit(itemId, index, textSpan);

    // Delete button
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'delete-option-btn';
    deleteBtn.textContent = '✕';
    deleteBtn.onclick = () => deleteOption(itemId, index);

    optionDiv.appendChild(reorderBtns);
    optionDiv.appendChild(textSpan);
    optionDiv.appendChild(editBtn);
    optionDiv.appendChild(deleteBtn);
    list.appendChild(optionDiv);
  });
}

// Move option (up or down) - update local UI only
function moveOption(itemId, fromIndex, toIndex) {
  const options = dropdownData[itemId];
  if (toIndex < 0 || toIndex >= options.length) return;

  // Update array order
  const [movedItem] = options.splice(fromIndex, 1);
  options.splice(toIndex, 0, movedItem);
  renderOptionsList(itemId);
  updateSaveButtonState(itemId);

  // Highlight moved item (use requestAnimationFrame to ensure DOM is updated)
  requestAnimationFrame(() => {
    const list = document.getElementById(`list-${itemId}`);
    if (list && list.children[toIndex]) {
      list.children[toIndex].classList.add('highlight');
      setTimeout(() => {
        if (list.children[toIndex]) {
          list.children[toIndex].classList.remove('highlight');
        }
      }, 800);
    }
  });
}

// Start editing - update local UI only
function startEdit(itemId, index, textSpan) {
  const oldValue = dropdownData[itemId][index];
  const input = document.createElement('input');
  input.id = `edit-${itemId}-${index}`;
  input.name = `edit-${itemId}-${index}`;
  input.type = 'text';
  input.className = 'edit-input';
  input.value = oldValue;

  const finishEdit = () => {
    const newValue = input.value.trim();
    if (newValue && newValue !== oldValue) {
      // Check for duplicates
      if (dropdownData[itemId].includes(newValue) && dropdownData[itemId].indexOf(newValue) !== index) {
        alert('此選項已存在');
        input.value = oldValue;
        return;
      }

      // Update local data
      dropdownData[itemId][index] = newValue;
      renderOptionsList(itemId);
      updateSaveButtonState(itemId);
    } else {
      renderOptionsList(itemId);
    }
  };

  input.onblur = finishEdit;
  input.onkeypress = (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      input.blur();
    }
  };
  input.onkeydown = (e) => {
    if (e.key === 'Escape') {
      renderOptionsList(itemId);
    }
  };

  textSpan.replaceWith(input);
  input.focus();
  input.select();
}

// Add option - update local UI only
function addOption(itemId) {
  const input = document.getElementById(`input-${itemId}`);
  const value = input.value.trim();

  if (!value) {
    alert('請輸入選項內容');
    return;
  }

  if (!dropdownData[itemId]) dropdownData[itemId] = [];

  if (dropdownData[itemId].includes(value)) {
    alert('此選項已存在');
    return;
  }

  // Update local data
  dropdownData[itemId].push(value);
  input.value = '';
  renderOptionsList(itemId);
  updateSaveButtonState(itemId);
}

// Delete option - update local UI only
function deleteOption(itemId, index) {
  const option = dropdownData[itemId][index];
  if (confirm(`確定要刪除「${option}」嗎？`)) {
    // Update local data
    dropdownData[itemId].splice(index, 1);
    renderOptionsList(itemId);
    updateSaveButtonState(itemId);
  }
}

// Check if there are unsaved changes
function hasChanges(itemId) {
  const current = dropdownData[itemId] || [];
  const original = originalDropdownData[itemId] || [];
  if (current.length !== original.length) return true;
  return current.some((val, idx) => val !== original[idx]);
}

// Update save button state
function updateSaveButtonState(itemId) {
  const saveBtn = document.getElementById(`save-btn-${itemId}`);
  const hintSpan = document.getElementById(`unsaved-hint-${itemId}`);
  const changed = hasChanges(itemId);

  hasUnsavedChanges[itemId] = changed;

  if (saveBtn) {
    saveBtn.disabled = !changed;
  }
  if (hintSpan) {
    hintSpan.textContent = changed ? '（有未儲存的變更）' : '';
  }
}

// Handle back button click
function handleBackClick(itemId) {
  if (hasUnsavedChanges[itemId]) {
    if (!confirm('您有未儲存的變更，確定要離開嗎？')) {
      return;
    }
    // Restore local data
    dropdownData[itemId] = [...originalDropdownData[itemId]];
    hasUnsavedChanges[itemId] = false;
  }
  showMenu();
}

// Save changes to backend
async function saveChanges(itemId) {
  const saveBtn = document.getElementById(`save-btn-${itemId}`);
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.textContent = '儲存中...';
  }

  // Show loading and disable all buttons
  showLoading();
  const allButtons = document.querySelectorAll('button, .settings-btn, .add-option-btn, .delete-option-btn');
  allButtons.forEach(btn => {
    btn.dataset._prevDisabled = btn.disabled ? '1' : '0';
    btn.disabled = true;
  });

  try {
    // Get original and current data
    const original = originalDropdownData[itemId] || [];
    const current = dropdownData[itemId] || [];

    // Send batch update request
    const col = columnMapping[itemId];
    const postData = {
      name: 'Batch Update Dropdown',
      itemId,
      action: 'batchUpdate',
      column: col,
      originalData: original,
      newData: current
    };

    // Only "category" needs to update both expense and budget sheets
    const isCategory = itemId === 'category';

    if (isCategory) {
      const [resExpense, resBudget] = await Promise.all([
        fetch(baseExpense, {
          method: "POST",
          redirect: "follow",
          mode: "cors",
          body: JSON.stringify(postData)
        }),
        fetch(baseBudget, {
          method: "POST",
          redirect: "follow",
          mode: "cors",
          body: JSON.stringify(postData)
        })
      ]);

      const resultExpense = await resExpense.json();
      const resultBudget = await resBudget.json();

      if (!resultExpense.success || !resultBudget.success) {
        throw new Error(`同步失敗：\n支出表：${resultExpense.message || '未知錯誤'}\n預算表：${resultBudget.message || '未知錯誤'}`);
      }
    } else {
      const resExpense = await fetch(baseExpense, {
        method: "POST",
        redirect: "follow",
        mode: "cors",
        body: JSON.stringify(postData)
      });

      const resultExpense = await resExpense.json();

      if (!resultExpense.success) {
        throw new Error(resultExpense.message || '同步失敗');
      }
    }

    // Update original data after success
    originalDropdownData[itemId] = [...current];
    hasUnsavedChanges[itemId] = false;

    // Notify other pages to reload dropdown
    try {
      const updateTime = Date.now().toString();
      localStorage.setItem('dropdownUpdated', updateTime);
      localStorage.setItem('dropdownUpdatedItemId', itemId);
      localStorage.setItem('dropdownUpdatedAction', 'batchUpdate');
      window.dispatchEvent(new StorageEvent('storage', {
        key: 'dropdownUpdated',
        newValue: updateTime
      }));
      window.dispatchEvent(new CustomEvent('dropdownUpdated', {
        detail: { itemId, action: 'batchUpdate', updateTime }
      }));
    } catch (e) {}

    alert('儲存成功！');

  } catch (error) {
    alert('儲存失敗：' + error.message);
  } finally {
    // Unlock
    hideLoading();
    allButtons.forEach(btn => {
      if (btn.dataset._prevDisabled === '0') {
        btn.disabled = false;
      }
      delete btn.dataset._prevDisabled;
    });

    if (saveBtn) {
      saveBtn.textContent = '修改';
      updateSaveButtonState(itemId);
    }
  }
}

// Show specified panel (show loading after clicking category, ensure data is loaded)
async function showPanel(panelId) {
  // If dropdown data not loaded yet, show loading and load once
  if (!dropdownLoaded) {
    showLoading();
    await loadDropdownData();
    hideLoading();
  }

  // Save original data for comparison
  originalDropdownData[panelId] = [...(dropdownData[panelId] || [])];
  hasUnsavedChanges[panelId] = false;

  document.getElementById('main-menu').style.display = 'none';
  document.querySelectorAll('.content-panel').forEach(p => p.classList.remove('active'));
  document.getElementById(`panel-${panelId}`).classList.add('active');
  renderOptionsList(panelId);
  updateSaveButtonState(panelId);
}

// Show main menu
function showMenu() {
  document.querySelectorAll('.content-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('main-menu').style.display = 'flex';
}

// Initialize page
async function init() {
  // When entering settings page, create main menu and category panels first for immediate interaction
  root.appendChild(createMainMenu());
  settingsItems.forEach(item => {
    root.appendChild(createPanel(item));
  });

  // Start loading dropdown sheet=1 data in background (non-blocking)
  loadDropdownData();
}

// Prevent bounce effect
(function() {
  let touchStartY = 0;
  let touchStartScrollTop = 0;
  let startScrollContainer = null;

  // Get scroll container
  function getScrollContainer(target) {
    const dropdownContainer = target.closest('.category-dropdown, .select-dropdown, .options-list, .select-options');
    if (dropdownContainer) {
      return dropdownContainer;
    }
    return document.documentElement;
  }

  document.addEventListener('touchstart', function(e) {
    touchStartY = e.touches[0].clientY;
    startScrollContainer = getScrollContainer(e.target);

    if (startScrollContainer === document.documentElement) {
      touchStartScrollTop = window.pageYOffset || document.documentElement.scrollTop;
    } else {
      touchStartScrollTop = startScrollContainer.scrollTop;
    }
  }, { passive: true });

  document.addEventListener('touchmove', function(e) {
    if (!startScrollContainer) {
      return;
    }

    const currentY = e.touches[0].clientY;
    const deltaY = currentY - touchStartY;

    let currentScrollTop, scrollHeight, clientHeight;

    if (startScrollContainer === document.documentElement) {
      currentScrollTop = window.pageYOffset || document.documentElement.scrollTop;
      scrollHeight = document.documentElement.scrollHeight;
      clientHeight = document.documentElement.clientHeight;
    } else {
      currentScrollTop = startScrollContainer.scrollTop;
      scrollHeight = startScrollContainer.scrollHeight;
      clientHeight = startScrollContainer.clientHeight;
    }

    const isAtTop = currentScrollTop <= 1;
    const isAtBottom = currentScrollTop + clientHeight >= scrollHeight - 1;

    // Prevent bounce at boundaries
    if (isAtTop && deltaY > 0) {
      // Pulling down at top
      if (e.cancelable) {
        e.preventDefault();
      }
      if (startScrollContainer === document.documentElement) {
        window.scrollTo(0, 0);
        document.documentElement.scrollTop = 0;
        document.body.scrollTop = 0;
      } else if (startScrollContainer) {
        startScrollContainer.scrollTop = 0;
      }
    } else if (isAtBottom && deltaY < 0) {
      // Pulling up at bottom
      if (e.cancelable) {
        e.preventDefault();
      }
      const maxScroll = Math.max(0, scrollHeight - clientHeight);
      if (startScrollContainer === document.documentElement) {
        window.scrollTo(0, maxScroll);
        document.documentElement.scrollTop = maxScroll;
        document.body.scrollTop = maxScroll;
      } else if (startScrollContainer) {
        startScrollContainer.scrollTop = maxScroll;
      }
    }
  }, { passive: false });

  // Add bounce prevention for all dropdown containers
  function setupDropdownPrevention() {
    const dropdowns = document.querySelectorAll('.category-dropdown, .select-dropdown, .options-list, .select-options');
    dropdowns.forEach(dropdown => {
      // Avoid duplicate binding
      if (dropdown.dataset.bouncePrevented) return;
      dropdown.dataset.bouncePrevented = 'true';

      let dropdownTouchStartY = 0;

      dropdown.addEventListener('touchstart', function(e) {
        dropdownTouchStartY = e.touches[0].clientY;
      }, { passive: true });

      dropdown.addEventListener('touchmove', function(e) {
        const currentY = e.touches[0].clientY;
        const deltaY = currentY - dropdownTouchStartY;
        const currentScrollTop = dropdown.scrollTop;
        const scrollHeight = dropdown.scrollHeight;
        const clientHeight = dropdown.clientHeight;

        const isAtTop = currentScrollTop <= 1;
        const isAtBottom = currentScrollTop + clientHeight >= scrollHeight - 1;

        if (isAtTop && deltaY > 0) {
          if (e.cancelable) {
            e.preventDefault();
          }
          dropdown.scrollTop = 0;
        } else if (isAtBottom && deltaY < 0) {
          if (e.cancelable) {
            e.preventDefault();
          }
          dropdown.scrollTop = Math.max(0, scrollHeight - clientHeight);
        }
      }, { passive: false });
    });
  }

  // Setup dropdowns after page load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setupDropdownPrevention);
  } else {
    setupDropdownPrevention();
  }

  // Watch for dynamically added dropdowns
  const observer = new MutationObserver(() => {
    setupDropdownPrevention();
  });
  observer.observe(document.body, { childList: true, subtree: true });

  document.addEventListener('touchend', function() {
    startScrollContainer = null;
  }, { passive: true });

  document.addEventListener('touchcancel', function() {
    startScrollContainer = null;
  }, { passive: true });
})();

init();
</script>
