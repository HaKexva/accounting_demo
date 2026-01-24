---
layout: page
title: 支出
permalink: /expense/
---

<link rel="stylesheet" href="{{ '/assets/common.css' | relative_url }}?v={{ site.time | date: '%s' }}">
<link rel="stylesheet" href="{{ '/assets/expense-table.css' | relative_url }}?v={{ site.time | date: '%s' }}">

<div id="user-info"></div>

<script>

// Expense / Budget GAS Web App URL (latest)
// Expense uses "new" URL, budget keeps original URL
const baseExpense = "https://script.google.com/macros/s/AKfycbxpBh0QVSVTjylhh9cj7JG9d6aJi7L7y6pQPW88EbAsNtcd5ckucLagH8XpSAGa8IZt/exec";
const baseBudget  = "https://script.google.com/macros/s/AKfycbxkOKU5YxZWP1XTCFCF7a62Ar71fUz4Qw7tjF3MvMGkLTt6QzzhGLnDsD7wVI_cgpAR/exec";
// Currently selected spreadsheet tab index (corresponds to getSheets()[index] in Apps Script, 2 represents the third tab)
let currentSheetIndex = 2;
// All month tab names obtained from Show Tab Name
// If first two are invalid items (blank sheet, dropdown), they will be filtered out
let sheetNames = [];
// Record whether original data includes first two invalid items, used for correct sheetIndex calculation
let hasInvalidFirstTwoSheets = false;

// Pre-loaded data for all months (key: sheetIndex, value: { data: {...}, total: [...] })
let allMonthsData = {}; // Store data and totals for each month

// Pure add new mode, no need to record index
let allRecords = []; // For history list
let filteredRecords = []; // Filtered records
let currentRecordIndex = 0; // Current record index

// Filter records by type (expense page only shows expenses)
const filterRecordsByType = (type) => {
  filteredRecords = allRecords.filter(r => r.type === type);
};

// ===== Dropdown Options (primarily from "dropdown" sheet=1, below are defaults/fallbacks) =====
let EXPENSE_CATEGORY_OPTIONS = [
  { value: '生活花費：食', text: '生活花費：食' },
  { value: '生活花費：衣與外貌', text: '生活花費：衣與外貌' },
  { value: '生活花費：住、居家裝修、衛生用品、次月繳納帳單', text: '生活花費：住、居家裝修、衛生用品、次月繳納帳單' },
  { value: '生活花費：行', text: '生活花費：行' },
  { value: '生活花費：育', text: '生活花費：育' },
  { value: '生活花費：樂', text: '生活花費：樂' },
  { value: '生活花費：健（醫療）', text: '生活花費：健（醫療）' },
  { value: '生活花費：帳單', text: '生活花費：帳單' },
  { value: '儲蓄：退休金、醫療預備金、過年紅包支出', text: '儲蓄：退休金、醫療預備金、過年紅包支出' },
  { value: '家人：過年紅包、紀念日', text: '家人：過年紅包、紀念日' }
];
let PAYMENT_METHOD_OPTIONS = [
  { value: '現金', text: '現金' },
  { value: '信用卡', text: '信用卡' },
  { value: '轉帳', text: '轉帳' },
  { value: '存款或儲值的支出：LINE BANK / 悠遊付 / mos card / 髮果 等', text: '存款或儲值的支出：LINE BANK / 悠遊付 / mos card / 髮果 等' }
];
let CREDIT_CARD_PAYMENT_OPTIONS = [
  { value: '分期付款', text: '分期付款' },
  { value: '一次支付', text: '一次支付' }
];
let MONTH_PAYMENT_OPTIONS = [
  { value: '本月支付', text: '本月支付' },
  { value: '次月支付', text: '次月支付' }
];
let PAYMENT_PLATFORM_OPTIONS = [
  { value: 'LINE BANK', text: 'LINE BANK' },
  { value: '悠遊付', text: '悠遊付' },
  { value: 'mos card', text: 'mos card' },
  { value: '髮果', text: '髮果' }
];

// ===== Payment Method Keyword Detection =====
// Detect if payment method is credit card type (contains keywords like "credit card")
const isCreditCardPayment = (paymentMethod) => {
  if (!paymentMethod) return false;
  const keywords = ['信用卡', '刷卡', 'credit card', 'creditcard'];
  const lowerPayment = paymentMethod.toLowerCase();
  return keywords.some(keyword => lowerPayment.includes(keyword.toLowerCase()));
};

// Detect if payment method is deposit or stored value type (contains keywords like "deposit", "stored value")
const isStoredValuePayment = (paymentMethod) => {
  if (!paymentMethod) return false;
  const keywords = ['存款', '儲值', '儲值的支出', '預付'];
  const lowerPayment = paymentMethod.toLowerCase();
  return keywords.some(keyword => lowerPayment.includes(keyword.toLowerCase()));
};

// ===== Using Shared Cache Module (SyncStatus) =====
// Use SyncStatus module's cache functionality (defined in assets/sync-status.js)
const getFromIDB = (key) => SyncStatus.getFromCache(key);
const setToIDB = (key, value) => SyncStatus.setToCache(key, value);
const getCacheTimestamp = (key) => SyncStatus.getCacheTimestamp(key);

// 背景同步：從 API 載入最新資料
const syncFromAPI = async () => {
  SyncStatus.startSync();
  try {
    // Load latest month list
    await loadMonthNames();
    setToIDB('sheetNames', sheetNames).catch(() => {});
    setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});

    // Check if need to create new month
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const currentMonthStr = `${year}${month}`;
    if (!sheetNames.includes(currentMonthStr)) {
      try {
        await callAPI({ name: "Create Tab" });
        await loadMonthNames();
        setToIDB('sheetNames', sheetNames).catch(() => {});
        setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});
      } catch (e) {
        // Creation failed, retry later
      }
    }

    // Recalculate current month index
    const closestSheetIndex = findClosestMonth();
    currentSheetIndex = closestSheetIndex;

    // Load latest data for current month
    const currentMonthData = await loadMonthData(currentSheetIndex);
    allMonthsData[currentSheetIndex] = currentMonthData;
    setToIDB(`monthData_${currentSheetIndex}`, currentMonthData).catch(() => {});

    // Update display (user may already be viewing page)
    processDataFromResponse(currentMonthData.data, true);
    updateTotalDisplay();

    // Load current month budget
    await loadBudgetForMonth(currentSheetIndex);
    updateTotalDisplay();
    setToIDB('budgetTotals', budgetTotals).catch(() => {});

    // Sync complete
    SyncStatus.endSync(true);

    // Background preload other months
    preloadAllMonthsData()
      .then(() => {
        updateTotalDisplay();
        Object.keys(allMonthsData).forEach(idx => {
          setToIDB(`monthData_${idx}`, allMonthsData[idx]).catch(() => {});
        });
      })
      .catch(() => {});

    // 背景預載其他月份的預算
    const budgetSheetIndices = sheetNames.map((name, idx) => idx + 2);
    const otherBudgetIndices = budgetSheetIndices.filter(sheetIndex => sheetIndex !== currentSheetIndex);
    const budgetPromises = otherBudgetIndices.map(sheetIndex => loadBudgetForMonth(sheetIndex).catch(() => {}));

    Promise.all(budgetPromises).then(() => {
      setToIDB('budgetTotals', budgetTotals).catch(() => {});
    });
  } catch (e) {
    SyncStatus.endSync(false);
  }
};

// Load latest options from "dropdown" sheet=1 (only one small API call, very fast)
async function loadDropdownOptions() {
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
      // Find first key whose value is an array
      for (const key in responseData) {
        if (Array.isArray(responseData[key]) && responseData[key].length > 0) {
          data = responseData[key];
          break;
        }
      }
    }

    if (!data || !Array.isArray(data) || data.length === 0) {
      return;
    }

    const headerRow = data[0];
    // Find corresponding column
    const colCategory   = findHeaderColumn(headerRow, ['消費類別', '類別']);
    const colPayment    = findHeaderColumn(headerRow, ['支付方式']);
    const colCreditCard = findHeaderColumn(headerRow, ['信用卡支付方式']);
    const colMonthPay   = findHeaderColumn(headerRow, ['本月／次月支付']);
    const colPlatform   = findHeaderColumn(headerRow, ['支付平台', '平台']);

    const readColumn = (col) => {
      const arr = [];
      if (col < 0) return arr;
      const seen = new Set();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const raw = row[col];
        if (raw === undefined || raw === null) continue;
        const val = raw.toString().trim();
        if (!val || seen.has(val)) continue;
        seen.add(val);
        arr.push({ value: val, text: val });
      }
      return arr;
    };

    if (colCategory >= 0) {
      EXPENSE_CATEGORY_OPTIONS = readColumn(colCategory);
    }
    if (colPayment >= 0) {
      PAYMENT_METHOD_OPTIONS = readColumn(colPayment);
    }
    if (colCreditCard >= 0) {
      CREDIT_CARD_PAYMENT_OPTIONS = readColumn(colCreditCard);
    }
    if (colMonthPay >= 0) {
      MONTH_PAYMENT_OPTIONS = readColumn(colMonthPay);
    }
    if (colPlatform >= 0) {
      PAYMENT_PLATFORM_OPTIONS = readColumn(colPlatform);
    }

  } catch (err) {
  }
}

// Update specified select + custom display based on new options (if exists)
function updateSelectOptions(selectId, options) {
  const select = document.getElementById(selectId);
  if (!select) {
    return;
  }

  if (!options || options.length === 0) {
    return;
  }

  const currentValue = select.value;
  select.innerHTML = '';

  options.forEach(opt => {
    const o = document.createElement('option');
    o.value = opt.value;
    o.textContent = opt.text;
    select.appendChild(o);
  });

  // Restore previously selected value (if still exists)
  if (currentValue && options.some(o => o.value === currentValue)) {
    select.value = currentValue;
  }

  // Update custom dropdown display text
  const container = select.parentElement;
  if (container) {
    const display = container.querySelector('.select-display');
    if (display) {
      const textEl = display.querySelector('.select-text');
      if (textEl) {
        const selectedOpt = select.options[select.selectedIndex];
        textEl.textContent = selectedOpt ? selectedOpt.textContent : '';
      }

      // Update dropdown option list
      const dropdown = container.querySelector('.select-dropdown');
      if (dropdown) {
        dropdown.innerHTML = '';
        options.forEach(opt => {
          const item = document.createElement('div');
          item.className = 'select-option';
          item.textContent = opt.text;
          item.onclick = () => {
            select.value = opt.value;
            if (textEl) textEl.textContent = opt.text;
            dropdown.style.display = 'none';
            const arrow = container.querySelector('.select-arrow');
            if (arrow) arrow.style.transform = 'rotate(0deg)';
            // Trigger change event
            select.dispatchEvent(new Event('change'));
          };
          dropdown.appendChild(item);
        });
      }
    }
  }
}

function findHeaderColumn(headerRow, keywords) {
  for (let c = 0; c < headerRow.length; c++) {
    const headerText = (headerRow[c] || '').toString().trim();
    if (!headerText) continue;
    if (keywords.some(k => headerText.includes(k))) return c;
  }
  return -1;
}

// Clear form, prepare for adding new entry
function clearForm() {
    const itemInput = document.getElementById('item-input');
  const expenseCategorySelect = document.getElementById('expense-category-select');
  const paymentMethodSelect = document.getElementById('payment-method-select');
  const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
  const monthPaymentSelect = document.getElementById('month-payment-select');
  const paymentPlatformSelect = document.getElementById('payment-platform-select');
  const actualCostInput = document.getElementById('actual-cost-input');
  const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');

  if (itemInput) itemInput.value = '';
  if (expenseCategorySelect && expenseCategorySelect.options.length > 0) {
    expenseCategorySelect.value = expenseCategorySelect.options[0].value;
    const selectContainer = expenseCategorySelect.parentElement;
        if (selectContainer) {
      const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
        const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
          selectText.textContent = expenseCategorySelect.options[0].textContent;
        }
      }
    }
  }
  // Payment method doesn't auto-reset, keep user's previous selection
  // if (paymentMethodSelect && paymentMethodSelect.options.length > 0) {
  //   paymentMethodSelect.value = paymentMethodSelect.options[0].value;
  //   const selectContainer = paymentMethodSelect.parentElement;
  //   if (selectContainer) {
  //     const selectDisplay = selectContainer.querySelector('.select-display');
  //     if (selectDisplay) {
  //       const selectText = selectDisplay.querySelector('.select-text');
  //       if (selectText) {
  //         selectText.textContent = paymentMethodSelect.options[0].textContent;
  //       }
  //     }
  //   }
  //   paymentMethodSelect.dispatchEvent(new Event('change'));
  // }
  if (creditCardPaymentSelect) creditCardPaymentSelect.value = '';
  if (monthPaymentSelect) monthPaymentSelect.value = '';
  if (paymentPlatformSelect) paymentPlatformSelect.value = '';
  if (actualCostInput) actualCostInput.value = '';
  if (recordCostInput) recordCostInput.value = '';
    if (noteInput) noteInput.value = '';

  // Date field removed
}

// Get current date and format as YYYY/MM/DD (without time)
function getNowFormattedDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// Convert various time string formats to YYYY/MM/DD (without time)
function formatRecordDateTime(raw) {
  if (!raw) return '';
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) {
    if (typeof raw === 'string' && raw.includes('/')) {
      return raw.split(' ')[0];
    }
    return raw;
  }
  const year = dt.getFullYear();
  const month = String(dt.getMonth() + 1).padStart(2, '0');
  const day = String(dt.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// Unified API call function
async function callAPI(postData) {
  const response = await fetch(baseExpense, {
    method: "POST",
    redirect: "follow",
    mode: "cors",
    keepalive: true,
    body: JSON.stringify(postData)
  });

  const responseText = await response.text();
  // Some GAS may return empty string; treat as success (no additional data)
  if (!responseText || responseText.trim() === '') {
    return { success: true, data: null, total: null };
  }

  let result;
  try {
    result = JSON.parse(responseText);
  } catch (e) {
    throw new Error('後端響應格式錯誤: ' + responseText.substring(0, 100));
  }

  if (!response.ok || !result.success) {
    throw new Error(result.message || result.error || '操作失敗');
  }

  // No longer clear cache - caller will use result.data to update cache
  // This avoids unnecessary reloads and improves performance

  return result;
}

// ===== History List Refresh (unified logic, avoid duplication) =====
function refreshHistoryList() {
  const historyModal = document.querySelector('.history-modal');
  if (!historyModal) return;

  const newRecords = loadHistoryListFromCache(currentSheetIndex);
  const listElement = historyModal.querySelector('.history-list');
  if (!listElement) return;

  listElement.innerHTML = '';
  const displayRecords = newRecords.filter(r => !isHeaderRecord(r));
  if (displayRecords.length === 0) {
    const emptyMsg = document.createElement('div');
    emptyMsg.textContent = '尚無歷史紀錄';
    emptyMsg.style.cssText = 'text-align: center; padding: 40px; color: #999;';
    listElement.appendChild(emptyMsg);
  } else {
    displayRecords.forEach(record => {
      listElement.appendChild(createHistoryItem(record, listElement));
    });
  }
}

// Unified update cache and history list
// If data and total are provided, use directly, otherwise reload
async function updateCacheAndHistory(resultData = null, resultTotal = null) {
  // If data is returned, use directly, otherwise reload
  if (resultData && resultTotal) {
    // Process returned data and update cache (backend returns data as array format)
    processDataFromResponse(resultData, false, currentSheetIndex);
    allMonthsData[currentSheetIndex] = { data: resultData, total: resultTotal };

    // Save to IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});
  } else {
    // No returned data, reload
    try {
      const monthData = await loadMonthData(currentSheetIndex);
      allMonthsData[currentSheetIndex] = monthData;
    } catch (error) {
    }
  }

  // Update history list
  refreshHistoryList();
}

// Fill form fields (for edit mode)
function fillForm(row) {
  if (!row) return;

  // Based on Apps Script structure: [time(0), item(1), category(2), spendWay(3), creditCard(4), monthIndex(5), actualCost(6), payment(7), recordCost(8), note(9)]
  setTimeout(() => {
    const itemInput = document.getElementById('item-input');
    const expenseCategorySelect = document.getElementById('expense-category-select');
    const paymentMethodSelect = document.getElementById('payment-method-select');
    const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
    const monthPaymentSelect = document.getElementById('month-payment-select');
    const paymentPlatformSelect = document.getElementById('payment-platform-select');
    const actualCostInput = document.getElementById('actual-cost-input');
    const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');

    // Item
    if (itemInput) {
      itemInput.value = row[1] || '';
    }

    // Category
    if (expenseCategorySelect) {
      expenseCategorySelect.value = row[2] || '';
      // Sync update custom dropdown display text
      const selectContainer = expenseCategorySelect.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
          if (selectText) {
            const selectedOption = expenseCategorySelect.options[expenseCategorySelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[2] || '';
          }
        }
      }
    }

    // Payment method
    if (paymentMethodSelect) {
      paymentMethodSelect.value = row[3] || '';
      // Sync update custom dropdown display text
      const selectContainer = paymentMethodSelect.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
          if (selectText) {
            const selectedOption = paymentMethodSelect.options[paymentMethodSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[3] || '';
          }
        }
      }
      // Trigger payment method change to show/hide related fields
      paymentMethodSelect.dispatchEvent(new Event('change'));
    }

    // Credit card payment method (if payment method is credit card type)
    if (creditCardPaymentSelect && isCreditCardPayment(row[3])) {
      creditCardPaymentSelect.value = row[4] || '';
      // Sync update custom dropdown display text
      const selectContainer = creditCardPaymentSelect.parentElement;
        if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
            const selectedOption = creditCardPaymentSelect.options[creditCardPaymentSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[4] || '';
          }
        }
      }
    }

    // This month/next month payment (if payment method is credit card type)
    if (monthPaymentSelect && isCreditCardPayment(row[3])) {
      monthPaymentSelect.value = row[5] || '';
      // Sync update custom dropdown display text
      const selectContainer = monthPaymentSelect.parentElement;
        if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
            const selectedOption = monthPaymentSelect.options[monthPaymentSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[5] || '';
          }
        }
      }
    }

    // Payment platform (if payment method is deposit or stored value type)
    if (paymentPlatformSelect && isStoredValuePayment(row[3])) {
      paymentPlatformSelect.value = row[7] || '';
      // Sync update custom dropdown display text
      const selectContainer = paymentPlatformSelect.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
          if (selectText) {
            const selectedOption = paymentPlatformSelect.options[paymentPlatformSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[7] || '';
          }
        }
      }
    }

    // Actual cost
    if (actualCostInput) {
      actualCostInput.value = row[6] || '';
    }

    // Record cost
    if (recordCostInput) {
      recordCostInput.value = row[8] || '';
    }

    // Note
    if (noteInput) {
      noteInput.value = row[9] || '';
    }

  }, 150);
}

// Check if a row is a header row (e.g., time/date, item, etc.), skip when displaying history
function isHeaderRecord(record) {
  if (!record || !record.row) return false;
  const row = record.row;
  const c0 = (row[0] || '').toString().trim();
  const c1 = (row[1] || '').toString().trim();
  if (!c0 && !c1) return false;
  const headerWords0 = ['時間', '日期'];
  const headerWords1 = ['項目', '品項', '標題'];
  const isHeader0 = headerWords0.some(w => c0.includes(w));
  const isHeader1 = headerWords1.some(w => c1.includes(w));
  return isHeader0 && isHeader1;
}

// ===== Budget Cache and Pre-aggregation (loaded from budget table) =====

// "Budget expense summary" for each month, key: sheetIndex, value: { [category: string]: number }
const budgetTotals = {};

// Load data for specified month from budget table, and pre-sum budget for each category
async function loadBudgetForMonth(sheetIndex) {
  // Return from cache if available
  if (budgetTotals[sheetIndex]) {
    return budgetTotals[sheetIndex];
  }

  const params = { name: "Show Tab Data", sheet: sheetIndex, _t: Date.now() };
  const url = `${baseBudget}?${new URLSearchParams(params)}`;
  let res;
  try {
    res = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store"
    });
  } catch (err) {
    throw new Error(`無法連接預算表伺服器: ${err.message}`);
  }

  if (!res.ok) {
    throw new Error(`載入預算資料失敗: HTTP ${res.status} ${res.statusText}`);
  }

  let data;
  try {
    data = await res.json();
  } catch (jsonErr) {
    const text = await res.text();
    throw new Error('預算表回應格式錯誤');
  }

  const categoryTotals = {};
  let processedRowsCount = 0;
  let skippedRowsCount = 0;
  let skippedRowsReasons = {
    notExpenseBudget: 0,
    emptyRow: 0,
    headerOrTotal: 0,
    noCategory: 0,
    invalidCost: 0
  };

  if (data && typeof data === 'object') {
    Object.keys(data).forEach(key => {
      const rows = data[key] || [];
      console.log(rows)
      // Only process named ranges containing "expense" (e.g., current month expense budget 202512)
      const isExpenseBudget = key.includes('支出');
      if (!isExpenseBudget) {
        skippedRowsReasons.notExpenseBudget++;
        return;
      }

      rows.forEach((row, rowIndex) => {
        if (!row || row.length === 0) {
          skippedRowsReasons.emptyRow++;
          return;
        }

        const firstCell = row[0];
        const firstCellStr = String(firstCell || '').trim();

        // Skip header rows and total rows
        // Budget format: [number, time, category, item, cost, note]
        // Header row's firstCell may be "number" or other header text
        // Total row's firstCell may be "total" or other total marker
        // Data row's firstCell is usually a number (number) or empty string
        if (firstCellStr === '編號' || firstCellStr === '總計' ||
            firstCellStr.toLowerCase() === '編號' || firstCellStr.toLowerCase() === '總計' ||
            firstCellStr.includes('編號') || firstCellStr.includes('總計')) {
          skippedRowsReasons.headerOrTotal++;
          return;
        }

        // If firstCell is a number, this is a valid data row (number)
        // If firstCell is empty string or null/undefined, it may also be a valid data row (allow empty number)
        // Continue processing

        // Real data row: allow firstCell to be empty, as long as there's category and amount
        // Category field: prefer column 3 (row[2]), if empty use column 2 (row[1]), to support "生活花費：食" format
        // Expense budget format: [number, time, category, item, cost, note]
        // Important: if category is the same, budgets should be summed, so the category's total budget is displayed correctly
        let category = (row[2] || '').toString().trim();
        if (!category) {
          category = (row[1] || '').toString().trim();
        }
        const item = (row[3] || '').toString().trim(); // item is in column 3 (index 3), only for logging

        const costRaw = row[4];
        const cost = parseFloat(costRaw);

        if (!category) {
          skippedRowsReasons.noCategory++;
          return;
        }
        if (!Number.isFinite(cost)) {
          skippedRowsReasons.invalidCost++;
          return;
        }

        // Important: if category is the same, budgets should be summed, not take the first one
        // Use category as key to ensure budgets with same category are summed
        // This way the expense table can correctly display the total budget for that category
        const budgetKey = category; // Only use category as key
        const oldTotal = categoryTotals[budgetKey] || 0;
        categoryTotals[budgetKey] = oldTotal + cost;
        processedRowsCount++;
      });
    });
  }

  budgetTotals[sheetIndex] = categoryTotals;

  return categoryTotals;
}

// Process data returned from Apps Script (for updating allRecords)
const processDataFromResponse = (data, shouldFilter = true, sheetIndexForContext = null) => {
  // Debug: log loaded data
  const effectiveIdx = (sheetIndexForContext !== undefined && sheetIndexForContext !== null) ? sheetIndexForContext : currentSheetIndex;
  const monthIdx = effectiveIdx - 2;
  const monthName = (monthIdx >= 0 && monthIdx < sheetNames.length) ? sheetNames[monthIdx] : 'unknown';
  console.log(`[Expense] Loading month: ${monthName} (sheetIndex: ${effectiveIdx})`);
  console.log('[Expense] Raw data:', data);
  
  // First clear current records
  allRecords = [];

  // If data is array format, need to convert first
  if (Array.isArray(data)) {
    const convertedData = {};
    let expenseRows = [];

    data.forEach((row, rowIndex) => {
      if (!row || row.length === 0) return;

      // Skip header row (first row is usually header)
      if (rowIndex === 0) {
        const firstCell = String(row[0] || '').trim().toLowerCase();
        if (firstCell === '交易日期' || firstCell === '時間' || firstCell === '日期' ||
            firstCell.includes('項目') || firstCell.includes('金額')) {
          return; // Skip header row
        }
      }

      // Skip total row
      const firstCell = String(row[0] || '').trim();
      if (firstCell === '總計' || firstCell === 'Total' || firstCell === '') {
        return;
      }

      // Based on Apps Script structure: [time(0), item(1), category(2), spendWay(3), creditCard(4), monthIndex(5), actualCost(6), payment(7), recordCost(8), note(9)]
      // Expense page processes all rows matching this structure (10 columns)
      if (row.length >= 10) {
        // Check if first column is time format or valid data
        const timeValue = row[0];
        // If first column is time format (contains / or -), or valid date string, or non-empty value, treat as valid record
        // Exclude obvious header rows (e.g., "交易日期")
        const timeStr = String(timeValue || '').trim();
        const isHeader = timeStr.toLowerCase() === '交易日期' ||
                         timeStr.toLowerCase() === '時間' ||
                         timeStr.toLowerCase() === '日期';

        if (!isHeader && timeValue !== null && timeValue !== undefined && timeStr !== '') {
          // Further check: if date format or non-empty string, add it
          if (timeStr.includes('/') || timeStr.includes('-') ||
              !isNaN(Date.parse(timeValue)) || timeStr.length > 0) {
          expenseRows.push(row);
          }
        }
      }
    });

    // key needs to include month name for subsequent filtering
    const effectiveSheetIndex = (sheetIndexForContext !== undefined && sheetIndexForContext !== null) ? sheetIndexForContext : currentSheetIndex;
    const monthIdx = effectiveSheetIndex - 2;
    const monthName = (monthIdx >= 0 && monthIdx < sheetNames.length) ? sheetNames[monthIdx] : '';
    if (expenseRows.length > 0) convertedData[`當月支出預算${monthName}`] = expenseRows;

    data = convertedData;
  }

  if (data && typeof data === 'object' && !Array.isArray(data)) {
    // Get current month name (allow override sheetIndex to avoid displaying wrong month)
    const effectiveSheetIndex = (sheetIndexForContext !== undefined && sheetIndexForContext !== null) ? sheetIndexForContext : currentSheetIndex;
    const monthIndex = effectiveSheetIndex - 2;
    const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length && sheetNames.length > 0) ? sheetNames[monthIndex] : '';

    Object.keys(data).forEach(key => {
      const rows = data[key] || [];

      // Determine type based on named range name
      // Named range format: current month income 202506 or current month expense budget 202506
      const isIncome = key.includes('收入');
      const isExpense = key.includes('支出');

      if (!isIncome && !isExpense) {
        return; // Skip data that is not income or expense
      }

      // Only process current month's data (named range name should include current month)
      // Fix: stricter month check to avoid mixing different months' data
      let isCurrentMonth = false;
      
      // If no month name (initialization stage), only accept exactly matched generic key
      if (!currentMonthName || currentMonthName === '') {
        // Only accept pure generic key (without month number)
        isCurrentMonth = key === '當月支出預算' || key === '當月收入';
      } else {
        // When month name exists, must exactly match current month
        // E.g., currentMonthName = '202602', key must be '當月支出預算202602'
        const exactMatchKey = key === `當月支出預算${currentMonthName}` || key === `當月收入${currentMonthName}`;
        const keyContainsMonth = key.includes(currentMonthName);
        isCurrentMonth = exactMatchKey || keyContainsMonth;
      }
      
      if (!isCurrentMonth) {
        return; // Skip data that is not current month
      }

      const type = isIncome ? '收入' : '支出';

      rows.forEach(row => {
        if (!row || row.length === 0) return;

        // Skip total rows and empty rows
        const firstCell = (row[0] || '').toString().trim();
        if (firstCell === '交易日期' || firstCell === '總計' || firstCell === 'Total' || firstCell === '') {
          return;
        }

        // For expense type, first column is time format, not number
        // So no need to check if it's a number, just ensure it's not a header row
        const isHeader = firstCell.toLowerCase() === '交易日期' ||
                         firstCell.toLowerCase() === '時間' ||
                         firstCell.toLowerCase() === '日期' ||
                         firstCell.includes('項目') ||
                         firstCell.includes('金額');
        if (isHeader) {
          return;
        }

        // Use time, item, and amount as unique identifier to avoid duplicate additions
        const rowKey = `${row[0] || ''}_${row[1] || ''}_${row[8] || ''}`;

        allRecords.push({ type, row });
      });
    });
  }

  // Debug: log processed data
  console.log(`[Expense] Processing complete, total ${allRecords.length} records`);
  if (allRecords.length > 0) {
    console.log('[Expense] Record list:', allRecords.map(r => ({
      type: r.type,
      時間: r.row[0],
      項目: r.row[1],
      類別: r.row[2],
      金額: r.row[8]
    })));
  }

  // Filter records by currently selected type (default shows expense)
  if (shouldFilter) {
    filterRecordsByType('支出'); // Now only expense

    // Expense page only has add new mode, no need to show history
    if (filteredRecords.length >= 1) {
      isNewMode = true;
      currentRecordIndex = 0;
    }
  }
};

// Recalculate top "Budget / Expense / Balance"
// Budget: read category's budget from budget table based on currently selected category
// Expense: sum category's expenses from all history records of current month based on currently selected category
const updateTotalDisplay = () => {
  const recordCostInput = document.getElementById('record-cost-input');
  const expenseCategorySelect = document.getElementById('expense-category-select');

  // Get currently selected category
  const selectedCategory = expenseCategorySelect ? expenseCategorySelect.value : '';
  // If no category selected, display 0
  if (!selectedCategory) {
    incomeAmount.textContent = '0';
    expenseAmount.textContent = '0';
    totalAmount.textContent = '0';
    updateTotalColor(0);
    return;
  }

  // 1. Budget: read current category's budget from budget table's budgetTotals
  let budget = 0;
  const budgetData = budgetTotals[currentSheetIndex];
  if (budgetData && typeof budgetData === 'object') {
    budget = parseFloat(budgetData[selectedCategory] || 0) || 0;
  }

  // 2. Expense: sum current category's expenses from all history records of current month
  let historyExpense = 0;
  let records = [];
  const monthData = allMonthsData[currentSheetIndex];
  if (monthData && monthData.data) {
    records = loadHistoryListFromCache(currentSheetIndex);
  }

  if (Array.isArray(records) && records.length > 0) {
    const processedRecordKeys = new Set();

    historyExpense = records.reduce((sum, r) => {
      const row = r.row || [];

      // Check if category matches (category is at index 2)
      const recordCategory = (row[2] || '').toString().trim();
      if (recordCategory !== selectedCategory) {
        return sum; // Skip records with different category
      }

      // Use record cost (index 8)
      const raw = row[8] !== undefined && row[8] !== null && row[8] !== '' ? row[8] : 0;
      const num = parseFloat(raw) || 0;

      // Use time, item, and amount as unique identifier to avoid duplicate calculation
      const recordKey = `${row[0] || ''}_${row[1] || ''}_${raw}`;
      if (processedRecordKeys.has(recordKey)) {
        return sum;
      }
      processedRecordKeys.add(recordKey);

      return sum + num;
    }, 0);
  }

  // Live input record cost (only calculate for currently selected category)
  let liveInput = 0;
  if (recordCostInput && recordCostInput.value) {
    // Check if currently input category matches selected category
    const currentInputCategory = expenseCategorySelect ? expenseCategorySelect.value : '';
    if (currentInputCategory === selectedCategory) {
      liveInput = parseFloat(recordCostInput.value) || 0;
    }
  }

  const expense = historyExpense + liveInput;

  // 3. Balance: budget - expense
  const remain = budget - expense;

  // Format display (use thousand separator)
  incomeAmount.textContent = budget.toLocaleString('zh-TW');
  expenseAmount.textContent = expense.toLocaleString('zh-TW');
  totalAmount.textContent = remain.toLocaleString('zh-TW');
  updateTotalColor(remain);

  // Hide summary area loading overlay
  const overlay = document.querySelector('.summary-loading-overlay');
  if (overlay) {
    overlay.remove();
  }
};

// Load single month's data and totals
const loadMonthData = async (sheetIndex) => {
  // Validate if sheetIndex is valid
  if (!Number.isFinite(sheetIndex) || sheetIndex < 2) {
    throw new Error(`Invalid sheet index: ${sheetIndex}`);
  }

  // Fetch "current month income / expense" etc. from spreadsheet - add timestamp to avoid cache
  const dataParams = { name: "Show Tab Data", sheet: sheetIndex, _t: Date.now() };
  const dataUrl = `${baseExpense}?${new URLSearchParams(dataParams)}`;
  let res;
  try {
    res = await fetch(dataUrl, {
    method: "GET",
    redirect: "follow",
    mode: "cors",
      cache: "no-store" // Force no cache
  });
  } catch (fetchError) {
    throw new Error(`無法連接到伺服器: ${fetchError.message}。請檢查網路連接或 CORS 設定。`);
  }

  if (!res.ok) {
    throw new Error(`載入資料失敗: HTTP ${res.status} ${res.statusText}`);
  }

  let data;
  try {
    data = await res.json();
  } catch (jsonError) {
    const text = await res.text();
    throw new Error(`伺服器回應格式錯誤: ${jsonError.message}`);
  }

  // If data is array format (Apps Script ShowTabData returns getValues()), need to convert to object format
  if (Array.isArray(data)) {
    // Convert array to object format to be compatible with processDataFromResponse
    // Assume array contains all data rows, we need to distinguish income and expense based on actual situation
    // Since expense page only processes expense data, we convert it to object format
    const convertedData = {};

    // Expense records are usually: number, time, category, item, amount, note, ...
    // We put all data under "current month expense" key
    // If there's income data, may need to distinguish based on column count or other identifiers
    let expenseRows = [];
    let incomeRows = [];

    data.forEach((row, index) => {
      if (!row || row.length === 0) return;

      // Skip header rows or empty rows
      const firstCell = row[0];
      if (firstCell === '' || firstCell === null || firstCell === undefined) return;

      // Skip "total" rows
      if (firstCell === '總計' || firstCell === 'Total' || firstCell.toString().trim() === '總計') {
        return;
      }

      // Based on Apps Script structure: [time(0), item(1), category(2), spendWay(3), creditCard(4), monthIndex(5), actualCost(6), payment(7), recordCost(8), note(9)]
      // Expense page mainly processes expenses, all rows treated as expenses (because structure is same)
      if (row.length >= 10) {
        // Ensure it's a valid data row (time field is not empty)
        if (firstCell !== '' && firstCell !== null && firstCell !== undefined) {
          expenseRows.push(row);
        }
      }
    });

    // Build object format, key needs to include month name for subsequent filtering
    const monthIndex = sheetIndex - 2;
    const monthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : '';
    if (expenseRows.length > 0) {
      // key includes month name, e.g., "current month expense budget 202601"
      convertedData[`當月支出預算${monthName}`] = expenseRows;
    }
    if (incomeRows.length > 0) {
      convertedData[`當月收入${monthName}`] = incomeRows;
    }
    data = convertedData; // Assign converted data back
  } else {
  // Check data length for each key
  Object.keys(data).forEach(key => {
    const rows = data[key] || [];
  });
  }

  // Calculate totals based on data (local calculation is faster and more reliable than API)
  let calculatedIncome = 0;
  let calculatedExpense = 0;
  let incomeCount = 0;
  let expenseCount = 0;

  // Use Set to track processed records, avoid duplicate calculation
  const processedIncomeRecords = new Set();
  const processedExpenseRecords = new Set();

  if (data && typeof data === 'object') {
    // Get current month name (from sheetNames, sheetIndex is actual sheet index, need to subtract 2)
    const monthIndex = sheetIndex - 2;
    const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length && sheetNames.length > 0) ? sheetNames[monthIndex] : '';

    Object.keys(data).forEach(key => {
      const rows = data[key] || [];

      // Determine type based on named range name
      // Named range format: current month income 202506 or current month expense budget 202506
      const isIncome = key.includes('收入');
      const isExpense = key.includes('支出');

      if (!isIncome && !isExpense) {
        return; // Skip data that is not income or expense
      }

      // Only process current month's data (named range name should include current month)
      const isCurrentMonth = currentMonthName && key.includes(currentMonthName);
      if (!isCurrentMonth) {
        return; // Skip data that is not current month
      }

      rows.forEach((row, rowIndex) => {
        if (!row || row.length === 0) return;

        // Check if it's an empty row (all fields are empty)
        const isEmptyRow = row.every(cell => cell === '' || cell === null || cell === undefined);
        if (isEmptyRow) {
          return;
        }

        // Skip total rows (first column is "總計" / "Total" or empty string)
        const firstCell = row[0];
        if (firstCell === '交易日期' || firstCell === '總計' || firstCell === 'Total' || firstCell === '' || firstCell === null || firstCell === undefined) {
          return;
        }

        // Check if it's a valid record (first column should be a number)
        const num = parseInt(firstCell, 10);
        if (!Number.isFinite(num) || num <= 0) {
          return;
        }

        // Check if this record has been processed (use number+time as unique identifier)
        const recordKey = `${num}_${row[1] || ''}`;
        if (isIncome) {
          if (processedIncomeRecords.has(recordKey)) {
            return;
          }
          processedIncomeRecords.add(recordKey);
        } else if (isExpense) {
          if (processedExpenseRecords.has(recordKey)) {
            return;
          }
          processedExpenseRecords.add(recordKey);
        }

        // Income: [number, time, item, cost, note] - cost is at index 3 (column D)
        // Expense: [number, time, category, item, cost, note] - cost is at index 4 (column K)
        const costIndex = isIncome ? 3 : (isExpense ? 4 : -1);
        if (costIndex >= 0 && row[costIndex] !== undefined && row[costIndex] !== null && row[costIndex] !== '') {
          const cost = parseFloat(row[costIndex]);
          if (Number.isFinite(cost) && cost !== 0) { // 允許負數，但不累加0
            if (isIncome) {
              calculatedIncome += cost;
              incomeCount++;
            } else if (isExpense) {
              calculatedExpense += cost;
              expenseCount++;
            }
          }
        }
      });

    });
  }

  const calculatedTotal = calculatedIncome - calculatedExpense;
  const calculatedTotalData = [calculatedIncome, calculatedExpense, calculatedTotal];
  // Use calculated totals instead of API returned totals
  return { data, total: calculatedTotalData };
};

// Preload data for all months
const preloadAllMonthsData = async () => {
  if (sheetNames.length === 0) return;

  // Calculate total number of months to load (exclude current month, already loaded)
  const monthsToLoad = sheetNames.filter((name, idx) => {
    const sheetIndex = idx + 2;
    return sheetIndex !== currentSheetIndex;
  });

  const totalMonths = monthsToLoad.length;
  if (totalMonths === 0) return; // If no months to preload, return directly

  let loadedCount = 0;
  const baseProgress = 2; // Already loaded month list (1) and current month (1)
  const totalProgress = sheetNames.length + 1; // Total progress = month list (1) + all months

  // Sort from new to old (assume month strings like 202512), send all requests in parallel
  // Put currently selected month first, rest sorted by year/month from new to old
  const allMonths = sheetNames.map((name, idx) => ({ name, sheetIndex: idx + 2 }));
  const current = allMonths.find(m => m.sheetIndex === currentSheetIndex);
  const others = allMonths.filter(m => m.sheetIndex !== currentSheetIndex)
    .sort((a, b) => (parseInt(b.name, 10) || 0) - (parseInt(a.name, 10) || 0));
  const orderedMonths = current ? [current, ...others] : others;

  const tasks = orderedMonths.map(async ({ name, sheetIndex }) => {
    if (sheetIndex === currentSheetIndex) {
      return null; // Skip current month, already loaded
    }

    // First check if pre-cached data already exists
    if (allMonthsData[sheetIndex]) {
      loadedCount++;
      updateProgress(baseProgress + loadedCount, totalProgress, '載入月份（從快取）');
      return { sheetIndex, name, success: true, fromCache: true };
    }

    try {
      const monthData = await loadMonthData(sheetIndex);
      allMonthsData[sheetIndex] = monthData;

      // Save to IndexedDB
      setToIDB(`monthData_${sheetIndex}`, monthData).catch(() => {});

      // Update progress bar (baseProgress + number of loaded months)
      loadedCount++;
      updateProgress(baseProgress + loadedCount, totalProgress, '載入月份');

      return { sheetIndex, name, success: true, fromCache: false };
    } catch (error) {
      // Update progress even if failed
      loadedCount++;
      updateProgress(baseProgress + loadedCount, totalProgress, '載入月份');

      // Loading failed, continue processing other months
      return { sheetIndex, name, success: false, error: error.message || error.toString() };
    }
  });

  const results = await Promise.all(tasks);

  // Update progress bar to 100%
  updateProgress(totalProgress, totalProgress, '載入完成');

  Object.keys(allMonthsData).forEach(key => {
    const monthData = allMonthsData[key];
    const total = Array.isArray(monthData.total) ? monthData.total : 'N/A';
  });

  // Delay a bit before hiding progress bar so user can see "loading complete" message
  // Note: hideSpinner already has 800ms delay to show "completed", so no extra delay needed here
  setTimeout(() => {
    hideSpinner();
  }, 100);
};

// Load current month's data from memory (no request sent)
const loadContentFromMemory = async () => {
  // First clear current records (ensure different months' data don't mix)
  allRecords = [];
  filteredRecords = [];
  currentRecordIndex = 0;

  // Validate if currentSheetIndex is valid
  if (!Number.isFinite(currentSheetIndex) || currentSheetIndex < 2) {
    currentSheetIndex = 2; // Default to third tab
  }

  // First read data from memory
  let monthData = allMonthsData[currentSheetIndex];

  // If not in memory, try reading from cache
  if (!monthData) {
    try {
      const storedData = await getFromIDB(`monthData_${currentSheetIndex}`);
      if (storedData) {
        monthData = storedData;
        // Explicitly mark as loaded from cache
        monthData._fromCache = true;
        // Also load into memory
        allMonthsData[currentSheetIndex] = monthData;
      }
    } catch (e) {
      // Cache may be unavailable or data corrupted, ignore error
    }
  }

  if (!monthData) {
    return false; // Indicates need to reload
  }

  // Confirm loaded data
  const totalPreview = Array.isArray(monthData.total) ? monthData.total : 'N/A';

  // Process data (will automatically filter and display records)
  processDataFromResponse(monthData.data, true);

  // Update total display (use budget table cache + current category expense)
  updateTotalDisplay();

  // Expense page only has add new mode
  isNewMode = true;
  currentRecordIndex = 0;

  return true; // Indicates successfully loaded from memory
};

// Load current month's data (prefer reading from memory, send request if not available)
const loadContent = async (forceReload = false) => {
  // If not forcing reload, try reading from memory first
  if (!forceReload && await loadContentFromMemory()) {
    return; // Successfully loaded from memory, return directly
  }

  // If no data in memory, or need to force reload, send request
  try {
    const monthData = await loadMonthData(currentSheetIndex);

    // Update data in memory
    allMonthsData[currentSheetIndex] = monthData;

    // Save to IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});

    // Process data
    processDataFromResponse(monthData.data);

    // Update total display (use budget table cache + current category expense)
    updateTotalDisplay();

    // Expense page only has add new mode
    isNewMode = true;
    currentRecordIndex = 0;
  } catch (error) {
    throw error;
  }
};


const loadTotal = async (forceRefresh = false) => {
  // Validate if currentSheetIndex is valid
  if (!Number.isFinite(currentSheetIndex) || currentSheetIndex < 2) {
    currentSheetIndex = 2; // Default to third tab
  }

  // Prefer reading from budget cache (budget table source), unless force refresh
  if (!forceRefresh && budgetTotals[currentSheetIndex]) {
    updateTotalDisplay();
    return;
  }

  // Send request to "budget table" and aggregate budgets of same category
  try {
    await loadBudgetForMonth(currentSheetIndex);
    updateTotalDisplay();
  } catch (error) {
    // Don't throw error, just log, to avoid affecting other features
  }
};

// Progress bar animation variables
let progressAnimationTimer = null;
let progressAnimationStartTime = null;
let progressAnimationTarget = 99; // Target percentage for auto animation (99%)

// Update progress bar (actual progress, will override auto animation)
const updateProgress = (current, total, text = '載入中...') => {
  const progressContainer = document.getElementById('loading-progress');
  if (!progressContainer) return;

  const percentage = total > 0 ? Math.min(99, Math.round((current / total) * 99)) : 0;
  const progressBar = progressContainer.querySelector('.progress-bar');

  if (progressBar) {
    progressBar.style.width = `${percentage}%`;
    progressAnimationTarget = percentage; // Update target, but won't exceed 99%
  }
};

// Start progress bar auto animation (from 0% to 99% in 10 seconds)
const startProgressAnimation = () => {
  // Clear previous animation
  if (progressAnimationTimer) {
    clearInterval(progressAnimationTimer);
    progressAnimationTimer = null;
  }

  const progressContainer = document.getElementById('loading-progress');
  if (!progressContainer) return;

  const progressBar = progressContainer.querySelector('.progress-bar');
  if (!progressBar) return;

  progressAnimationStartTime = Date.now();
  const duration = 10000; // 10 seconds
  const startPercentage = 0;
  const endPercentage = 99;

  progressAnimationTimer = setInterval(() => {
    const elapsed = Date.now() - progressAnimationStartTime;
    const progress = Math.min(elapsed / duration, 1);
    const currentPercentage = startPercentage + (endPercentage - startPercentage) * progress;

    // Ensure doesn't exceed actual progress target
    const finalPercentage = Math.min(currentPercentage, progressAnimationTarget);

    progressBar.style.width = `${finalPercentage}%`;

    // If already reached 99% or exceeded, stop animation
    if (progress >= 1 || finalPercentage >= 99) {
      clearInterval(progressAnimationTimer);
      progressAnimationTimer = null;
    }
  }, 16); // ~60fps
};

const showSpinner = (coverHeader = false) => {
  // If already exists, remove first
  const existingOverlay = document.getElementById('loading-overlay');
  if (existingOverlay) {
    existingOverlay.remove();
  }

  // Check if history modal is open (z-index: 10000)
  const historyModal = document.querySelector('.history-modal');
  const isHistoryModalOpen = historyModal && window.getComputedStyle(historyModal).display !== 'none';

  // Decide whether to cover header
  const shouldCoverHeader = coverHeader || isHistoryModalOpen;
  // z-index: when covering header should be above header (2000), when modal open should be above modal (10000)
  const zIndexValue = isHistoryModalOpen ? 10001 : (coverHeader ? 2001 : 1500);

  // Create fullscreen overlay
  const overlay = document.createElement('div');
  overlay.id = 'loading-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: ${shouldCoverHeader ? '0' : '60px'};
    left: 0;
    right: 0;
    bottom: 0;
    background: linear-gradient(135deg, rgba(255, 245, 240, 0.97) 0%, rgba(255, 248, 245, 0.97) 50%, rgba(255, 240, 235, 0.97) 100%);
    z-index: ${zIndexValue};
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    cursor: not-allowed;
  `;

  const progressContainer = document.createElement('div');
  progressContainer.id = 'loading-progress';
  const wrapperDiv = document.createElement('div');
  wrapperDiv.style.cssText = 'width: 300px; text-align: center;';
  const progressText = document.createElement('div');
  progressText.textContent = '載入中...';
  progressText.style.cssText = 'font-size: 16px; margin-bottom: 15px; color: #333;';
  const bgDiv = document.createElement('div');
  bgDiv.style.cssText = 'width: 100%; height: 8px; background-color: #e0e0e0; border-radius: 4px; overflow: hidden;';
  const progressBar = document.createElement('div');
  progressBar.className = 'progress-bar';
  progressBar.style.width = '0%';
  bgDiv.appendChild(progressBar);
  wrapperDiv.appendChild(progressText);
  wrapperDiv.appendChild(bgDiv);
  progressContainer.appendChild(wrapperDiv);

  overlay.appendChild(progressContainer);
  document.body.appendChild(overlay);

  // Start auto progress bar animation (from 0% to 99% in 10 seconds)
  setTimeout(() => {
    startProgressAnimation();
  }, 50);
};

const hideSpinner = () => {
  // Clear auto animation timer
  if (progressAnimationTimer) {
    clearInterval(progressAnimationTimer);
    progressAnimationTimer = null;
  }

  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    const progressContainer = document.getElementById('loading-progress');
    if (progressContainer) {
      const progressBar = progressContainer.querySelector('.progress-bar');
      // Jump to 100% first, then remove
      if (progressBar) {
        progressBar.style.width = '100%';
      }
    }

    // Remove overlay after slight delay so user can see 100%
    setTimeout(() => {
      overlay.remove();
      const spinner = document.getElementById('loading-spinner');
      if (spinner) {
        spinner.remove();
      }
      const progress = document.getElementById('loading-progress');
      if (progress) {
        progress.remove();
      }
      // Reset target percentage
      progressAnimationTarget = 99;
    }, 200);
  } else {
    const spinner = document.getElementById('loading-spinner');
    if (spinner) {
      spinner.remove();
    }
    const progress = document.getElementById('loading-progress');
    if (progress) {
      progress.remove();
    }
  }
};

const createInputRow = (labelText, inputId, inputType = 'text') => {
  const row = document.createElement('div');
  row.className = 'input-row';

  const label = document.createElement('label');
  label.textContent = labelText;
  label.htmlFor = inputId; // Associate with input

  const input = document.createElement('input');
  input.id = inputId;
  input.name = inputId; // Add name attribute to support autofill
  input.type = inputType;

  row.appendChild(label);
  row.appendChild(input);
  return row;
};

const createTextareaRow = (labelText, textareaId, rows = 3) => {
  const row = document.createElement('div');
  row.className = 'input-row';

  const label = document.createElement('label');
  label.textContent = labelText;
  label.htmlFor = textareaId; // Associate with textarea

  const textarea = document.createElement('textarea');
  textarea.id = textareaId;
  textarea.name = textareaId; // Add name attribute to support autofill
  textarea.rows = rows;

  row.appendChild(label);
  row.appendChild(textarea);
  return row;
};

const createSelectRow = (labelText, selectId, options) => {
  const row = document.createElement('div');
  row.className = 'select-row';

  const label = document.createElement('label');
  label.textContent = labelText;
  label.htmlFor = selectId; // Associate with select

  const selectContainer = document.createElement('div');
  selectContainer.className = 'select-container';

  const selectDisplay = document.createElement('div');
  selectDisplay.className = 'select-display';

  // Handle empty options case
  const safeOptions = (options && options.length > 0) ? options : [{ value: '', text: '無選項' }];

  const selectText = document.createElement('div');
  selectText.className = 'select-text';
  selectText.textContent = safeOptions[0].text;

  const selectArrow = document.createElement('div');
  selectArrow.className = 'select-arrow';
  selectArrow.textContent = '▼';

  selectDisplay.appendChild(selectText);
  selectDisplay.appendChild(selectArrow);

  const hiddenSelect = document.createElement('select');
  hiddenSelect.id = selectId;
  hiddenSelect.name = selectId; // Add name attribute to support autofill
  hiddenSelect.style.display = 'none';
  hiddenSelect.value = safeOptions[0].value;

  const dropdown = document.createElement('div');
  dropdown.className = 'select-dropdown';

  safeOptions.forEach(opt => {
    const option = document.createElement('div');
    option.className = 'select-option';
    option.textContent = opt.text;
    option.dataset.value = opt.value;

    option.addEventListener('click', function() {
      selectText.textContent = opt.text;
      hiddenSelect.value = opt.value;
      dropdown.style.display = 'none';
      selectArrow.style.transform = 'rotate(0deg)';
      hiddenSelect.dispatchEvent(new Event('change'));
    });

    dropdown.appendChild(option);
    const hiddenOption = document.createElement('option');
    hiddenOption.value = opt.value;
    hiddenOption.textContent = opt.text;
    hiddenSelect.appendChild(hiddenOption);
  });

  selectDisplay.addEventListener('click', function(e) {
    e.stopPropagation();
    const isOpen = dropdown.style.display === 'block';

    // Close all other dropdowns
    document.querySelectorAll('.select-dropdown').forEach(otherDropdown => {
      if (otherDropdown !== dropdown) {
        otherDropdown.style.display = 'none';
        const otherContainer = otherDropdown.closest('.select-container');
        if (otherContainer) {
          const otherArrow = otherContainer.querySelector('.select-arrow');
          if (otherArrow) {
            otherArrow.style.transform = 'rotate(0deg)';
          }
        }
      }
    });

    dropdown.style.display = isOpen ? 'none' : 'block';
    selectArrow.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(180deg)';
  });

  document.addEventListener('click', function(e) {
    if (!selectContainer.contains(e.target)) {
      dropdown.style.display = 'none';
      selectArrow.style.transform = 'rotate(0deg)';
    }
  });

  selectContainer.appendChild(selectDisplay);
  selectContainer.appendChild(dropdown);
  selectContainer.appendChild(hiddenSelect);

  row.appendChild(label);
  row.appendChild(selectContainer);
  return row;
};

// (Moved above: callAPI and updateCacheAndHistory)

// Get form data
function getFormData(prefix = '') {
  const itemInput = document.getElementById(prefix ? `${prefix}-item-input` : 'item-input');
  const dateInput = document.getElementById(prefix ? `${prefix}-date-input` : 'date-input');
  const expenseCategorySelect = document.getElementById(prefix ? `${prefix}-expense-category-select` : 'expense-category-select');
  const paymentMethodSelect = document.getElementById(prefix ? `${prefix}-payment-method-select` : 'payment-method-select');
  const creditCardPaymentSelect = document.getElementById(prefix ? `${prefix}-credit-card-payment-select` : 'credit-card-payment-select');
  const monthPaymentSelect = document.getElementById(prefix ? `${prefix}-month-payment-select` : 'month-payment-select');
  const paymentPlatformSelect = document.getElementById(prefix ? `${prefix}-payment-platform-select` : 'payment-platform-select');
  const actualCostInput = document.getElementById(prefix ? `${prefix}-actual-cost-input` : 'actual-cost-input');
  const recordCostInput = document.getElementById(prefix ? `${prefix}-record-cost-input` : 'record-cost-input');
  const noteInput = document.getElementById(prefix ? `${prefix}-note-input` : 'note-input');

  if (!itemInput || !expenseCategorySelect || !paymentMethodSelect || !actualCostInput || !recordCostInput) {
    throw new Error('請等待表單載入完成');
  }

  const item = itemInput.value.trim();
  // Get date and convert to YYYY/MM/DD format
  let date = '';
  if (dateInput && dateInput.value) {
    const dateValue = dateInput.value; // Format: YYYY-MM-DD
    const [year, month, day] = dateValue.split('-');
    date = `${year}/${month}/${day}`;
  } else {
    // If no date selected, use today
    date = getNowFormattedDateTime();
  }

  const category = expenseCategorySelect.value;
  const spendWay = paymentMethodSelect.value;
  const actualCostValue = actualCostInput.value.trim();
  const recordCostValue = recordCostInput.value.trim();
  const note = noteInput ? noteInput.value.trim() : '';

  if (!item) throw new Error('請輸入項目');
  if (!category) throw new Error('請選擇類別');
  if (!spendWay) throw new Error('請選擇支付方式');

  // 允許 0 值，但確保它們是有效數字（不能為負數）
  let actualCost = 0;
  let recordCost = 0;
  
  if (actualCostValue && actualCostValue.trim() !== '') {
    actualCost = parseFloat(actualCostValue);
    if (isNaN(actualCost) || actualCost < 0) {
      throw new Error('請輸入有效實際消費金額（不能為負數）');
    }
  }
  
  if (recordCostValue && recordCostValue.trim() !== '') {
    recordCost = parseFloat(recordCostValue);
    if (isNaN(recordCost) || recordCost < 0) {
      throw new Error('請輸入有效列帳消費金額（不能為負數）');
    }
  }
  
  // 兩者都可以是 0 或空（將預設為 0）

  let creditCard = '';
  let monthIndex = '';
  let payment = '';

  if (isCreditCardPayment(spendWay)) {
    creditCard = creditCardPaymentSelect ? creditCardPaymentSelect.value : '';
    // 取得月份支付值 - monthPaymentSelect 應該是來自 createSelectRow 的隱藏 <select> 元素
    // 隱藏的 select 元素具有 ID 'month-payment-select'（或帶前綴）
    const monthSelectElement = monthPaymentSelect || document.getElementById(prefix ? `${prefix}-month-payment-select` : 'month-payment-select');
    if (monthSelectElement && monthSelectElement.tagName === 'SELECT') {
      monthIndex = monthSelectElement.value || '';
      // 如果值為空，使用預設值（第一個選項）
      if (!monthIndex && monthSelectElement.options && monthSelectElement.options.length > 0) {
        monthIndex = monthSelectElement.options[0].value || '';
      }
    } else if (monthSelectElement) {
      // 如果它不是 SELECT 元素，嘗試在內部找到隱藏的 select
      const hiddenSelect = monthSelectElement.querySelector ? monthSelectElement.querySelector('select') : null;
      if (hiddenSelect) {
        monthIndex = hiddenSelect.value || '';
        // 如果值為空，使用預設值（第一個選項）
        if (!monthIndex && hiddenSelect.options && hiddenSelect.options.length > 0) {
          monthIndex = hiddenSelect.options[0].value || '';
        }
      }
    }
  } else if (isStoredValuePayment(spendWay)) {
    payment = paymentPlatformSelect ? paymentPlatformSelect.value : '';
  }

  return { date, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note };
}

const updateDivVisibility = (forceType = null) => {
  // 如果提供了類型參數，使用它；否則嘗試從 DOM 獲取最新元素的值
  let categoryValue = forceType;
  if (categoryValue === null) {
    // 先嘗試從全局變數獲取
    if (typeof categorySelect !== 'undefined' && categorySelect.value) {
      categoryValue = categorySelect.value;
    } else {
      // 如果全局變數不可用，從 DOM 獲取最新元素
      const categorySelectElement = document.getElementById('category-select');
      if (categorySelectElement) {
        categoryValue = categorySelectElement.value;
      } else {
        categoryValue = '支出'; // 默認值
      }
    }
  }

  div2.innerHTML = '';
  div3.innerHTML = '';
  div4.innerHTML = '';

  // 添加日期輸入字段（在所有類型中都顯示）
  const dateRow = createInputRow('日期：', 'date-input', 'date');
  const dateInput = dateRow.querySelector('#date-input');
  if (dateInput) {
    // 設置年份限制為四位數（1000-9999年）
    dateInput.min = '1000-01-01';
    dateInput.max = '9999-12-31';
    // 添加事件監聽器來驗證年份為四位數
    dateInput.addEventListener('input', (e) => {
      const value = e.target.value;
      if (value) {
        const year = parseInt(value.split('-')[0]);
        if (!isNaN(year) && (year < 1000 || year > 9999)) {
          const parts = value.split('-');
          if (parts[0].length > 4) {
            parts[0] = parts[0].substring(0, 4);
            const correctedValue = parts.join('-');
            if (/^\d{4}-\d{2}-\d{2}/.test(correctedValue)) {
              e.target.value = correctedValue;
            }
          }
        }
      }
    });
    // 設置默認值為今天
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    dateInput.value = `${year}-${month}-${day}`;
  }

  if (categoryValue === '支出') {
    const categoryRow = createSelectRow('類別：', 'expense-category-select', [
      { value: '生活花費：食', text: '生活花費：食' },
      { value: '生活花費：衣與外貌', text: '生活花費：衣與外貌' },
      { value: '生活花費：住、居家裝修、衛生用品、次月繳納帳單', text: '生活花費：住、居家裝修、衛生用品、次月繳納帳單' },
      { value: '生活花費：行', text: '生活花費：行' },
      { value: '生活花費：育', text: '生活花費：育' },
      { value: '生活花費：樂', text: '生活花費：樂' },
      { value: '生活花費：健（醫療）', text: '生活花費：健（醫療）' },
      { value: '生活花費：帳單', text: '生活花費：帳單' },
      { value: '儲蓄：退休金、醫療預備金、過年紅包支出', text: '儲蓄：退休金、醫療預備金、過年紅包支出' },
      { value: '家人：過年紅包、紀念日', text: '家人：過年紅包、紀念日' }
    ]);
    const costRow = createInputRow('金額：', 'cost-input', 'number');
    const noteRow = createTextareaRow('備註：', 'note-input', 3);
    noteRow.style.marginBottom = '0px';

    div2.appendChild(dateRow);
    div2.appendChild(categoryRow);
    div3.appendChild(costRow);
    div4.appendChild(noteRow);

    itemContainer.style.display = 'flex';
    div2.style.display = 'flex';
    div3.style.display = 'flex';
    div4.style.display = 'flex';
  } else if (categoryValue === '收入') {
    const costRow = createInputRow('金額：', 'cost-input', 'number');
    const noteRow = createTextareaRow('備註：', 'note-input', 3);
    noteRow.style.marginBottom = '0px';

    div2.appendChild(dateRow);
    div2.appendChild(costRow);
    div3.appendChild(noteRow);

    itemContainer.style.display = 'flex';
    div2.style.display = 'flex';
    div3.style.display = 'flex';
    div4.style.display = 'none';
  }
};

const saveData = async () => {
  // Lock entire page, wait for backend response
  showSpinner();
  mainSaveButton.textContent = '儲存中...';
  mainSaveButton.disabled = true;
  mainSaveButton.style.opacity = '0.6';
  mainSaveButton.style.cursor = 'not-allowed';

  // Disable all inputs and buttons
  const itemInput = document.getElementById('item-input');
  const expenseCategorySelect = document.getElementById('expense-category-select');
  const paymentMethodSelect = document.getElementById('payment-method-select');
  const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
  const monthPaymentSelect = document.getElementById('month-payment-select');
  const paymentPlatformSelect = document.getElementById('payment-platform-select');
  const actualCostInput = document.getElementById('actual-cost-input');
  const recordCostInput = document.getElementById('record-cost-input');
  const noteInput = document.getElementById('note-input');
  if (itemInput) itemInput.disabled = true;
  if (expenseCategorySelect) expenseCategorySelect.disabled = true;
  if (paymentMethodSelect) paymentMethodSelect.disabled = true;
  if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = true;
  if (monthPaymentSelect) monthPaymentSelect.disabled = true;
  if (paymentPlatformSelect) paymentPlatformSelect.disabled = true;
  if (actualCostInput) actualCostInput.disabled = true;
  if (recordCostInput) recordCostInput.disabled = true;
  if (noteInput) noteInput.disabled = true;
  if (historyButton) historyButton.disabled = true;

  try {
    // 首先取得支付方式以檢查是否為信用卡
    const paymentMethodValue = paymentMethodSelect ? paymentMethodSelect.value : '';
    
    // 如果是信用卡，確保正確取得月份支付選擇值
    let monthValue = '';
    if (isCreditCardPayment(paymentMethodValue)) {
      // 直接取得隱藏的 select 元素值
      const monthSelectElement = document.getElementById('month-payment-select');
      if (monthSelectElement && monthSelectElement.tagName === 'SELECT') {
        monthValue = monthSelectElement.value || '';
        console.log('[Expense] Directly getting month payment value:', {
          element: monthSelectElement,
          value: monthValue,
          selectedIndex: monthSelectElement.selectedIndex,
          options: Array.from(monthSelectElement.options).map(o => ({ value: o.value, text: o.text, selected: o.selected }))
        });
      } else {
        console.warn('[Expense] month-payment-select element not found or not a SELECT:', monthSelectElement);
      }
      
      if (!monthValue || monthValue.trim() === '') {
        alert('請選擇「本月支付」或「次月支付」');
        // 恢復按鈕狀態
        mainSaveButton.textContent = '儲存';
        mainSaveButton.disabled = false;
        mainSaveButton.style.opacity = '1';
        mainSaveButton.style.cursor = 'pointer';
        // 恢復所有輸入
        if (itemInput) itemInput.disabled = false;
        if (expenseCategorySelect) expenseCategorySelect.disabled = false;
        if (paymentMethodSelect) paymentMethodSelect.disabled = false;
        if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
        if (monthPaymentSelect) monthPaymentSelect.disabled = false;
        if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
        if (actualCostInput) actualCostInput.disabled = false;
        if (recordCostInput) recordCostInput.disabled = false;
        if (noteInput) noteInput.disabled = false;
        if (historyButton) historyButton.disabled = false;
        hideSpinner();
        return;
      }
    }
    
    const formData = getFormData();
    const monthIndex = currentSheetIndex - 2;
    const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : '';
    
    // 如果是信用卡支付，用直接取得的值覆蓋 monthIndex
    if (isCreditCardPayment(formData.spendWay) && monthValue) {
      formData.monthIndex = monthValue;
    }
    
    // 將 monthIndex 轉換為 month 供後端 API 使用（後端期望 'month' 而非 'monthIndex'）
    const apiData = {
      name: "Upsert Data",
      sheet: currentSheetIndex,
      date: formData.date,
      item: formData.item,
      category: formData.category,
      spendWay: formData.spendWay,
      creditCard: formData.creditCard,
      month: formData.monthIndex || '', // 後端期望 'month' 但前端使用 'monthIndex'
      actualCost: formData.actualCost,
      payment: formData.payment,
      recordCost: formData.recordCost,
      note: formData.note
    };
    
    // 調試日誌：檢查發送的資料
    if (isCreditCardPayment(formData.spendWay)) {
      console.log('[Expense] Saving credit card expense:', {
        spendWay: formData.spendWay,
        creditCard: formData.creditCard,
        monthIndex: formData.monthIndex,
        month: apiData.month,
        monthValue: monthValue,
        apiData: apiData
      });
    }
    
    const result = await callAPI(apiData);

    // Only show success message after backend response
    alert('資料已成功儲存！');

    // Update cache with returned data
    if (result && result.data) {
      allMonthsData[currentSheetIndex] = { data: result.data, total: result.total };
      processDataFromResponse(result.data, false, currentSheetIndex);
      // Sync refresh history list
      refreshHistoryList();

      // Save to IndexedDB
      setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});
    }

    // Update total display (force reload from API)
    await loadTotal(true);

    // After save complete, clear form to prepare for next entry
    clearForm();
  } catch (error) {
    alert('儲存失敗: ' + error.message);
  } finally {
    // Restore all buttons and inputs
    hideSpinner();
    mainSaveButton.textContent = '儲存';
    mainSaveButton.disabled = false;
    mainSaveButton.style.opacity = '1';
    mainSaveButton.style.cursor = 'pointer';

    if (itemInput) itemInput.disabled = false;
    if (expenseCategorySelect) expenseCategorySelect.disabled = false;
    if (paymentMethodSelect) paymentMethodSelect.disabled = false;
    if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
    if (monthPaymentSelect) monthPaymentSelect.disabled = false;
    if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
    if (actualCostInput) actualCostInput.disabled = false;
    if (recordCostInput) recordCostInput.disabled = false;
    if (noteInput) noteInput.disabled = false;
    if (historyButton) historyButton.disabled = false;
  }
};


const totalContainer = document.createElement('div');
totalContainer.className = 'total-container';

const budgetCardsContainer = document.createElement('div');
budgetCardsContainer.className = 'budget-cards-container';

// 日期已移除


// 刪除功能已移除（純新增模式不需要）
// 刪除功能改為在歷史紀錄彈出視窗中處理

// 歷史紀錄按鈕（時鐘圖標）
// 歷史紀錄按鈕（文字按鈕）
const historyButton = document.createElement('button');
historyButton.className = 'history-button';
historyButton.textContent = '歷史紀錄';
historyButton.style.cssText = `
  display: inline-block;
  margin-left: 20px;
  padding: 4px 12px;
  border: 1px solid #ddd;
  border-radius: 4px;
  background-color: #fff;
  color: #333;
  font-size: 14px;
  cursor: pointer;
  transition: all 0.2s;
`;
historyButton.onmouseenter = () => {
  historyButton.style.backgroundColor = '#f5f5f5';
};
historyButton.onmouseleave = () => {
  historyButton.style.backgroundColor = '#fff';
};

// 防止重複點擊的標誌
let isHistoryModalOpen = false;

// 月份選擇容器
let monthSelectContainer = null;

// 載入月份列表
async function loadMonthNames() {
  try {
    const params = { name: "Show Tab Name" };
    const url = `${baseExpense}?${new URLSearchParams(params)}&_t=${Date.now()}`;
    const response = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`載入失敗: HTTP ${response.status}`);
    }

    const data = await response.json();

    if (Array.isArray(data)) {
      // 檢查前兩個項目是否是無效項目（空白表、下拉選單）
      const invalidItems = ['空白表', '下拉選單', '', null, undefined];
      const firstItem = data[0];
      const secondItem = data[1];

      const shouldSkipFirst = invalidItems.includes(firstItem);
      const shouldSkipSecond = invalidItems.includes(secondItem);

      if (shouldSkipFirst && shouldSkipSecond && data.length > 2) {
        // 如果前兩個都是無效項目，跳過它們
        sheetNames = data.slice(2);
        hasInvalidFirstTwoSheets = true;
      } else {
        // 如果前兩個不是無效項目，使用所有數據（都是有效的月份）
        sheetNames = data;
        hasInvalidFirstTwoSheets = false;
      }

      // 保存到 IndexedDB
      setToIDB('sheetNames', sheetNames).catch(() => {});
      setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});

      return sheetNames;
    } else {
      return [];
    }
  } catch (error) {
    return [];
  }
}

// 顯示月份選擇
async function showMonthSelect() {
  // 如果已經打開，不重複打開
  if (isHistoryModalOpen) {
    return;
  }

  isHistoryModalOpen = true;
  historyButton.disabled = true;

  // 如果已經存在月份選擇容器，先移除
  if (monthSelectContainer) {
    monthSelectContainer.remove();
  }

  // 每次打開都重新載入月份列表，確保數據是最新的
  const months = await loadMonthNames();
  // 如果載入失敗或為空，顯示錯誤信息
  if (!months || months.length === 0) {
    alert('無法載入月份列表，請稍後再試');
    isHistoryModalOpen = false;
    historyButton.disabled = false;
    return;
  }

  // 創建彈出視窗
  const modal = document.createElement('div');
  modal.className = 'month-select-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 10000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;

  // 創建內容容器
  const content = document.createElement('div');
  content.className = 'month-select-modal-content';
  content.style.cssText = `
    background: linear-gradient(135deg, #fff8f5 0%, #fffaf8 50%, #fff5f0 100%);
    border-radius: 12px;
    padding: 30px;
    max-width: 400px;
    width: 100%;
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  // 關閉按鈕（右上角）
  const closeBtn = document.createElement('button');
  closeBtn.innerHTML = '✕';
  closeBtn.style.cssText = `
    position: absolute;
    top: 10px;
    right: 10px;
    width: 30px;
    height: 30px;
    border: none;
    background: transparent;
    font-size: 24px;
    cursor: pointer;
    color: #666;
    line-height: 1;
  `;
  closeBtn.onclick = () => {
    modal.remove();
    isHistoryModalOpen = false;
    historyButton.disabled = false;
  };

  // 標題
  const title = document.createElement('h2');
  title.textContent = '選擇月份';
  title.style.cssText = `
    margin: 0 0 20px 0;
    font-size: 20px;
    font-weight: 600;
  `;

  // 創建下拉選單容器
  const selectContainer = document.createElement('div');
  selectContainer.style.cssText = `
    margin-bottom: 20px;
  `;

  // 創建下拉選單
  const select = document.createElement('select');
  select.id = 'month-select';
  select.name = 'month-select';
  select.style.cssText = `
    width: 100%;
    padding: 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 16px;
    cursor: pointer;
  `;

  // 添加選項
  if (months.length === 0) {
    const option = document.createElement('option');
    option.value = '';
    option.textContent = '載入中...';
    option.disabled = true;
    select.appendChild(option);
  } else {
    months.forEach((month, index) => {
      const option = document.createElement('option');
      // 修正：直接使用陣列索引計算 sheetIndex
      // sheetNames 已跳過前兩個無效項目（空白表、下拉選單）
      // 所以 sheetIndex = index + 2
      const sheetIndex = index + 2;

      option.value = sheetIndex;
      option.textContent = month;
      option.dataset.monthName = month; // 儲存月份名稱以便調試
      if (sheetIndex === currentSheetIndex) {
        option.selected = true;
      }
      select.appendChild(option);
    });
  }

  // 選擇月份後載入該月份的歷史紀錄
  select.addEventListener('change', async (e) => {
    const selectedSheetIndex = parseInt(e.target.value, 10);
    if (isNaN(selectedSheetIndex)) {
      return;
    }

    const selectedOption = e.target.options[e.target.selectedIndex];
    const selectedMonthName = selectedOption ? selectedOption.dataset.monthName || selectedOption.textContent : '';
    currentSheetIndex = selectedSheetIndex;

      // Close month selection modal
      modal.remove();
      isHistoryModalOpen = false;

      // Show history modal
      await showHistoryModal();
    });

  selectContainer.appendChild(select);

  content.appendChild(closeBtn);
  content.appendChild(title);
  content.appendChild(selectContainer);
  modal.appendChild(content);

  // Click background to close
  modal.onclick = (e) => {
    if (e.target === modal) {
      modal.remove();
      isHistoryModalOpen = false;
      historyButton.disabled = false;
    }
  };

  document.body.appendChild(modal);
}

historyButton.onclick = showHistoryModal;

// Helper function to create history item
function createHistoryItem(record, listElement) {
  const item = document.createElement('div');
  item.className = 'history-item';
  item.style.cssText = `
    padding: 12px;
    border-bottom: 1px solid #eee;
    cursor: pointer;
    transition: background-color 0.2s;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: relative;
  `;
  item.onmouseenter = () => item.style.backgroundColor = '#f5f5f5';
  item.onmouseleave = () => item.style.backgroundColor = 'transparent';

  const itemContent = document.createElement('div');
  itemContent.style.cssText = 'flex: 1; min-width: 0;';

  const date = document.createElement('div');
  date.textContent = formatRecordDateTime(record.row[0] || '');
  date.style.cssText = 'font-size: 14px; color: #666; margin-bottom: 4px;';

  const itemTitle = document.createElement('div');
  itemTitle.textContent = record.row[1] || '(無標題)';
  itemTitle.style.cssText = 'font-size: 16px; font-weight: 500; color: #333; margin-bottom: 4px;';

  // Add category display (category - row[2])
  const category = document.createElement('div');
  const categoryText = (record.row[2] || '').trim();
  category.textContent = categoryText ? `類別：${categoryText}` : '';
  category.style.cssText = 'font-size: 14px; color: #666; margin-bottom: 4px;';

  // Add cost display (record cost - recordCost)
  const cost = document.createElement('div');
  const recordCost = parseFloat(record.row[8]) || 0;
  cost.textContent = `金額：${recordCost.toLocaleString('zh-TW', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
  cost.style.cssText = 'font-size: 14px; color: #e74c3c; font-weight: 500;';

  itemContent.appendChild(date);
  itemContent.appendChild(itemTitle);
  if (categoryText) {
    itemContent.appendChild(category);
  }
  itemContent.appendChild(cost);

  // Delete button
  const deleteBtn = document.createElement('button');
  deleteBtn.textContent = '刪除';
  deleteBtn.style.cssText = `
    padding: 4px 12px;
    border: 1px solid #dc3545;
    border-radius: 4px;
    background-color: #fff;
    color: #dc3545;
    font-size: 12px;
    cursor: pointer;
    margin-left: 10px;
    transition: all 0.2s;
    flex-shrink: 0;
  `;
  deleteBtn.onmouseenter = () => {
    deleteBtn.style.backgroundColor = '#dc3545';
    deleteBtn.style.color = '#fff';
  };
  deleteBtn.onmouseleave = () => {
    deleteBtn.style.backgroundColor = '#fff';
    deleteBtn.style.color = '#dc3545';
  };
  deleteBtn.onclick = async (e) => {
    e.stopPropagation();
    if (!confirm('確定要刪除這筆記錄嗎？')) return;

    // Lock entire page, wait for backend response
    showSpinner();
    deleteBtn.disabled = true;

    // Disable all inputs and buttons
    const itemInput = document.getElementById('item-input');
    const expenseCategorySelect = document.getElementById('expense-category-select');
    const paymentMethodSelect = document.getElementById('payment-method-select');
    const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
    const monthPaymentSelect = document.getElementById('month-payment-select');
    const paymentPlatformSelect = document.getElementById('payment-platform-select');
    const actualCostInput = document.getElementById('actual-cost-input');
    const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.disabled = true;
    if (expenseCategorySelect) expenseCategorySelect.disabled = true;
    if (paymentMethodSelect) paymentMethodSelect.disabled = true;
    if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = true;
    if (monthPaymentSelect) monthPaymentSelect.disabled = true;
    if (paymentPlatformSelect) paymentPlatformSelect.disabled = true;
    if (actualCostInput) actualCostInput.disabled = true;
    if (recordCostInput) recordCostInput.disabled = true;
    if (noteInput) noteInput.disabled = true;
    if (mainSaveButton) mainSaveButton.disabled = true;
    if (historyButton) historyButton.disabled = true;

    try {
      await deleteRecord(record);
      // Only show success message after backend response
      alert('記錄已成功刪除！');
      // Remove this record from display
      if (listElement && item.parentNode === listElement) {
        listElement.removeChild(item);
      } else {
        item.remove();
      }
    } catch (error) {
      alert('刪除失敗: ' + error.message);
    } finally {
      // Restore all buttons and inputs
      hideSpinner();
      deleteBtn.disabled = false;
      if (itemInput) itemInput.disabled = false;
      if (expenseCategorySelect) expenseCategorySelect.disabled = false;
      if (paymentMethodSelect) paymentMethodSelect.disabled = false;
      if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
      if (monthPaymentSelect) monthPaymentSelect.disabled = false;
      if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
      if (actualCostInput) actualCostInput.disabled = false;
      if (recordCostInput) recordCostInput.disabled = false;
      if (noteInput) noteInput.disabled = false;
      if (mainSaveButton) mainSaveButton.disabled = false;
      if (historyButton) historyButton.disabled = false;
    }
  };

  item.appendChild(itemContent);
  item.appendChild(deleteBtn);

  item.onclick = (e) => {
    if (e.target === deleteBtn || deleteBtn.contains(e.target)) {
      return;
    }
    showEditModal(record);
  };

  return item;
}

// Load history list from cache (no request sent, only read from memory)
function loadHistoryListFromCache(sheetIndex) {
  // First clear current records
  allRecords = [];

  // Read data from memory (don't use IndexedDB to avoid async complexity)
  let monthData = allMonthsData[sheetIndex];

  if (!monthData || !monthData.data) {
    return [];
  }

  // Process data (no filtering, only need record list)
  // Pass sheetIndex to ensure correct processing of month data
  processDataFromResponse(monthData.data, false, sheetIndex);

  // Annotate each record with corresponding spreadsheet row number
  // Note: Since processDataFromResponse already filtered header and total rows,
  // we use simple index calculation (row 1 is header, so start from row 2)
  // If data is array format, ShowTabData returns data where first row (index 0) is header, data starts from second row (index 1)
  // But since we already filtered header row, index here corresponds to filtered record index
  // Actual sheet row number needs to consider: header row (row 1) + filtered total rows
  // For simplicity, we assume records are in order, use index + 2 (row 1 is header, row 2 onwards is data)
  allRecords.forEach((r, index) => {
    r.sheetRowIndex = index + 2;
  });
  return allRecords;
}

// Find closest month (current month or latest month)
// Fix: directly search from sheetNames array to avoid errors from hardcoded reference points
function findClosestMonth() {
  // Select closest month based on current year/month
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const currentMonthStr = `${currentYear}${String(currentMonth).padStart(2, '0')}`;

  // If sheetNames has data, search from array
  if (sheetNames.length > 0) {
    // Try to find current month
    const currentIndex = sheetNames.findIndex(name => name === currentMonthStr);
    if (currentIndex !== -1) {
      // sheetNames already skipped first two invalid items (blank sheet, dropdown)
      // So actual sheet index is currentIndex + 2
      return currentIndex + 2;
    }

    // If current month doesn't exist, find closest month (prefer latest)
    // Convert month strings to numbers for comparison
    const monthNumbers = sheetNames.map(name => parseInt(name, 10)).filter(n => !isNaN(n));
    const currentMonthNum = parseInt(currentMonthStr, 10);

    // Find maximum value less than or equal to current month (closest past or current month)
    let closestMonth = null;
    let closestIndex = -1;
    for (let i = 0; i < sheetNames.length; i++) {
      const monthNum = parseInt(sheetNames[i], 10);
      if (!isNaN(monthNum) && monthNum <= currentMonthNum) {
        if (closestMonth === null || monthNum > closestMonth) {
          closestMonth = monthNum;
          closestIndex = i;
        }
      }
    }

    // If found, return corresponding sheet index
    if (closestIndex !== -1) {
      return closestIndex + 2;
    }

    // If not found, return last (latest) month
    return sheetNames.length - 1 + 2;
  }

  // If sheetNames has no data, return default value
  return 2;
}

// Show history modal (includes month selection)
async function showHistoryModal() {
  // Check if history modal already exists
  const existingModal = document.querySelector('.history-modal');
  if (existingModal && isHistoryModalOpen) {
    return;
  }

  // If existing modal exists, remove it first
  if (existingModal) {
    existingModal.remove();
  }

  isHistoryModalOpen = true;
  historyButton.disabled = true; // Disable button to prevent repeated clicks

  // Create modal (show first, prepare to show loading)
  const modal = document.createElement('div');
  modal.className = 'history-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 10000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;

  // Create content container
  const content = document.createElement('div');
  content.className = 'history-modal-content';
  content.style.cssText = `
    background: linear-gradient(135deg, #fff8f5 0%, #fffaf8 50%, #fff5f0 100%);
    border-radius: 12px;
    padding: 0 20px 20px 20px; /* Top edge aligned, remove top padding */
    max-width: 600px;
    width: 100%;
    max-height: 80vh;
    overflow-y: auto;
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  // Set content container min height to ensure loading display is correct
  content.style.minHeight = '300px';

  // Create loading display (using CSS spinner)
  const loadingDiv = document.createElement('div');
  loadingDiv.id = 'history-loading-spinner';
  loadingDiv.style.cssText = `
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 250px;
    width: 100%;
  `;

  const spinnerIcon = document.createElement('div');
  spinnerIcon.style.cssText = `
    width: 48px;
    height: 48px;
    border: 4px solid #e0e0e0;
    border-top-color: #3498db;
    border-radius: 50%;
    animation: spin 1s linear infinite;
    margin-bottom: 20px;
  `;

  const loadingText = document.createElement('div');
  loadingText.textContent = '載入中...';
  loadingText.style.cssText = `
    font-size: 16px;
    color: #666;
  `;

  loadingDiv.appendChild(spinnerIcon);
  loadingDiv.appendChild(loadingText);

  // Find closest month
  const closestSheetIndex = findClosestMonth();
  currentSheetIndex = closestSheetIndex;

  // First check if cached data exists (memory or IndexedDB)
  let hasCachedData = false;
  if (!allMonthsData[currentSheetIndex]) {
    try {
      const storedData = await getFromIDB(`monthData_${currentSheetIndex}`);
      if (storedData) {
        allMonthsData[currentSheetIndex] = storedData;
        hasCachedData = true;
      }
    } catch (e) {}
  } else {
    hasCachedData = true;
  }

  // If cached, display directly (don't show spinner)
  // If not cached, show spinner and load data
  if (!hasCachedData) {
    content.appendChild(loadingDiv);
  }
  modal.appendChild(content);
  document.body.appendChild(modal);

  // If no cache, need to load data
  if (!hasCachedData) {
    // If month list not loaded yet, load it first
    if (sheetNames.length === 0) {
      await loadMonthNames();
    }

    try {
      const monthData = await loadMonthData(currentSheetIndex);
      allMonthsData[currentSheetIndex] = monthData;

      // Save to IndexedDB
      setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});
    } catch (e) {
    }

    // Remove spinner
    loadingDiv.remove();
  }

  // Load history list from cache
  const records = loadHistoryListFromCache(currentSheetIndex);

  // Mobile version add right padding
  if (window.innerWidth <= 768) {
    content.style.paddingRight = '20px';
    content.style.paddingLeft = '20px';
  }

  // Close button (placed in header, same row as title, both sticky)
  const closeBtn = document.createElement('button');
  closeBtn.innerHTML = '✕';
  closeBtn.style.cssText = `
    width: 28px;
    height: 28px;
    border: none;
    background: transparent;
    font-size: 20px;
    cursor: pointer;
    color: #666;
    line-height: 1;
  `;
  closeBtn.onclick = () => {
    modal.remove();
    isHistoryModalOpen = false;
    historyButton.disabled = false; // 重新啟用按鈕
  };

  // 篩選和排序狀態（在函數作用域內，需要保存當前的記錄引用）
  let currentFilterRecords = records; // 保存當前用於篩選的記錄
  let filterType = 'none'; // 'none', 'date', 'category'
  let filterValue = '';
  let sortType = 'none'; // 'none', 'date_asc', 'date_desc', 'category'

  // 將日期轉換為本地日期的 YYYY-MM-DD 格式（避免時區偏移問題）
  // 使用與 formatRecordDateTime 相同的邏輯，但返回 YYYY-MM-DD 格式以便與篩選值比較
  const formatDateToLocalString = (date) => {
    if (!date && date !== 0) return '';
    
    // 如果是Date對象，直接使用本地時間格式化（與 formatRecordDateTime 相同邏輯）
    if (date instanceof Date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    
    const dateStr = String(date).trim();
    
    // 如果已經是 YYYY-MM-DD 格式（不含時間），直接返回
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      return dateStr;
    }
    
    // 如果是 YYYY-MM-DD 格式但包含時間部分，只取日期部分
    const dateOnlyMatch = dateStr.match(/^(\d{4}-\d{2}-\d{2})/);
    if (dateOnlyMatch) {
      return dateOnlyMatch[1];
    }
    
    // 如果是 YYYY/MM/DD 格式，轉換為 YYYY-MM-DD
    if (dateStr.includes('/')) {
      const parts = dateStr.split('/');
      if (parts.length >= 3) {
        const year = parts[0].trim();
        const month = parts[1].trim().padStart(2, '0');
        const day = parts[2].trim().split(' ')[0].split('T')[0].padStart(2, '0');
        // 驗證年份是四位數
        if (year.length === 4 && /^\d{4}$/.test(year)) {
          return `${year}-${month}-${day}`;
        }
      }
    }
    
    // 使用與 formatRecordDateTime 相同的邏輯來解析日期
    // 這是關鍵：確保篩選時使用的日期解析邏輯與顯示時一致
    const dt = new Date(dateStr);
    if (!Number.isNaN(dt.getTime())) {
      // 使用本地時間方法（與 formatRecordDateTime 相同）
      const year = dt.getFullYear();
      const month = String(dt.getMonth() + 1).padStart(2, '0');
      const day = String(dt.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    
    // 如果是數字，可能是時間戳
    if (typeof date === 'number') {
      const dateObj = new Date(date);
      if (!isNaN(dateObj.getTime())) {
        const year = dateObj.getFullYear();
        if (year >= 1900 && year <= 2100) {
          const month = String(dateObj.getMonth() + 1).padStart(2, '0');
          const day = String(dateObj.getDate()).padStart(2, '0');
          return `${year}-${month}-${day}`;
        }
      }
    }
    
    return '';
  };

  // 更新記錄列表顯示的函數（支持篩選）
  const updateHistoryList = (listElement, recordsToShow) => {
    listElement.innerHTML = '';

    // 顯示所有記錄（過濾標題行）
    let displayRecords = recordsToShow.filter(r => !isHeaderRecord(r));

    // 應用篩選
    if (filterType === 'date' && filterValue) {
      // 計算前一天和後一天的日期（處理時區偏移問題）
      const filterDate = new Date(filterValue + 'T00:00:00'); // 使用本地時區的午夜
      const prevDate = new Date(filterDate);
      prevDate.setDate(prevDate.getDate() - 1);
      const nextDate = new Date(filterDate);
      nextDate.setDate(nextDate.getDate() + 1);
      
      // 轉換為 YYYY-MM-DD 格式
      const formatDateOnly = (d) => {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      };
      
      const filterValueStr = formatDateOnly(filterDate);
      const prevDateStr = formatDateOnly(prevDate);
      const nextDateStr = formatDateOnly(nextDate);
      
      // 允許匹配篩選日期、前一天或後一天的記錄（處理時區偏移）
      displayRecords = displayRecords.filter(record => {
        const recordDate = record.row[0] || '';
        const recordDateStr = formatDateToLocalString(recordDate);
        // 匹配篩選日期、前一天或後一天
        return recordDateStr === filterValueStr || 
               recordDateStr === prevDateStr || 
               recordDateStr === nextDateStr;
      });
    } else if (filterType === 'category' && filterValue) {
      displayRecords = displayRecords.filter(record => {
        const recordCategory = (record.row[2] || '').trim();
        return recordCategory === filterValue;
      });
    }

    // 應用排序
    if (sortType === 'date_asc' || sortType === 'date_desc') {
      displayRecords.sort((a, b) => {
        const dateA = a.row[0] || '';
        const dateB = b.row[0] || '';
        
        // 轉換為可比較的日期字符串（使用本地時間）
        const dateStrA = formatDateToLocalString(dateA);
        const dateStrB = formatDateToLocalString(dateB);
        
        // 根據排序類型決定順序
        if (sortType === 'date_desc') {
          // 由大到小（從新到舊）
          return dateStrB.localeCompare(dateStrA);
        } else {
          // 由小到大（從舊到新）
          return dateStrA.localeCompare(dateStrB);
        }
      });
    } else if (sortType === 'category') {
      // 按照 EXPENSE_CATEGORY_OPTIONS 的順序排序
      displayRecords.sort((a, b) => {
        const categoryA = (a.row[2] || '').trim();
        const categoryB = (b.row[2] || '').trim();
        
        // 找到類別在選單中的索引
        const indexA = EXPENSE_CATEGORY_OPTIONS.findIndex(cat => cat.value === categoryA);
        const indexB = EXPENSE_CATEGORY_OPTIONS.findIndex(cat => cat.value === categoryB);
        
        // 如果類別不在選單中，放在最後
        if (indexA === -1 && indexB === -1) {
          return categoryA.localeCompare(categoryB, 'zh-TW');
        }
        if (indexA === -1) return 1; // A 不在選單中，放在後面
        if (indexB === -1) return -1; // B 不在選單中，放在後面
        
        // 按照選單順序排序
        return indexA - indexB;
      });
    }

    if (displayRecords.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.textContent = filterType !== 'none' ? '沒有符合篩選條件的記錄' : '尚無歷史紀錄';
      emptyMsg.style.cssText = 'text-align: center; padding: 40px; color: #999;';
      listElement.appendChild(emptyMsg);
  } else {
      displayRecords.forEach((record, index) => {
        const item = document.createElement('div');
        item.className = 'history-item';
        item.style.cssText = `
          padding: 12px;
          padding-right: ${window.innerWidth <= 768 ? '20px' : '12px'};
          border-bottom: 1px solid #eee;
          cursor: pointer;
          transition: background-color 0.2s;
          display: flex;
          align-items: center;
          justify-content: space-between;
          position: relative;
        `;
        listElement.appendChild(createHistoryItem(record, listElement));
      });
    }
  };

  // Title and month selection container (sticky, includes filters)
  const headerContainer = document.createElement('div');
  headerContainer.style.cssText = `
    display: flex;
    flex-direction: column;
    margin-bottom: 20px;
    position: sticky;
    top: 0;
    background: linear-gradient(135deg, #fff8f5 0%, #fffaf8 50%, #fff5f0 100%);
    padding: 15px 20px;
    margin-left: -20px;
    margin-right: -20px;
    z-index: 10;
    border-bottom: 1px solid #eee;
  `;
  
  // First row: title and close button
  const headerTopRow = document.createElement('div');
  headerTopRow.style.cssText = `
    display: flex;
    align-items: center;
    justify-content: space-between;
    width: 100%;
  `;

  const title = document.createElement('h2');
  title.textContent = '歷史紀錄';
  title.id = 'history-modal-title';
  title.style.cssText = `
    margin: 0;
    font-size: 20px;
    font-weight: 600;
  `;

  // Month selection dropdown
  monthSelectContainer = document.createElement('div');
  monthSelectContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 10px;
  `;

  const monthLabel = document.createElement('label');
  monthLabel.textContent = '月份：';
  monthLabel.htmlFor = 'history-month-select'; // Associate with select
  monthLabel.style.cssText = 'font-size: 14px; color: #666;';

  const monthSelect = document.createElement('select');
  monthSelect.id = 'history-month-select';
  monthSelect.name = 'history-month-select';
  monthSelect.style.cssText = `
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 14px;
    cursor: pointer;
  `;

  // Add month options
  sheetNames.forEach((month, index) => {
    const option = document.createElement('option');
    const sheetIndex = index + 2;
    option.value = sheetIndex;
    option.textContent = month;
    option.dataset.monthName = month; // Store month name for debugging
    if (sheetIndex === currentSheetIndex) {
      option.selected = true;
    }
    monthSelect.appendChild(option);
  });

  // Month selection change event
  monthSelect.addEventListener('change', async (e) => {
    const selectedSheetIndex = parseInt(e.target.value, 10);
    if (isNaN(selectedSheetIndex)) {
      return;
    }

    const selectedOption = e.target.options[e.target.selectedIndex];
    const selectedMonthName = selectedOption ? selectedOption.dataset.monthName || selectedOption.textContent : '';
    currentSheetIndex = selectedSheetIndex;

    // If that month's data not loaded yet, show loading
    if (!allMonthsData[currentSheetIndex]) {
      // Create loading display (using CSS spinner)
      const loadingDiv = document.createElement('div');
      loadingDiv.style.cssText = `
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 300px;
        width: 100%;
      `;

      const spinnerIcon = document.createElement('div');
      spinnerIcon.style.cssText = `
        width: 48px;
        height: 48px;
        border: 4px solid #e0e0e0;
        border-top-color: #3498db;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin-bottom: 20px;
      `;

      const loadingText = document.createElement('div');
      loadingText.textContent = '載入中...';
      loadingText.style.cssText = `
        font-size: 16px;
        color: #666;
      `;

      loadingDiv.appendChild(spinnerIcon);
      loadingDiv.appendChild(loadingText);
      list.innerHTML = '';
      list.appendChild(loadingDiv);

      // Load that month's data
      try {
        const monthData = await loadMonthData(currentSheetIndex);
        allMonthsData[currentSheetIndex] = monthData;

        // Save to IndexedDB
        setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});

        // Remove loading display
        loadingDiv.remove();
      } catch (e) {
        // Load failed, show error message
        list.innerHTML = '<div style="text-align: center; padding: 40px; color: #999;">載入失敗，請稍後再試</div>';
        return;
      }
    }

    // Reload that month's records from cache
    const newRecords = loadHistoryListFromCache(currentSheetIndex);
    
    // Update current record reference
    currentFilterRecords = newRecords;
    
    // Reset filter and update filter options (using new records)
    filterType = 'none';
    filterValue = '';
    filterTypeSelect.value = 'none';
    sortType = 'none';
    sortSelect.value = 'none';
    updateFilterValueOptions(newRecords);

    // Update record list display
    updateHistoryList(list, newRecords);
  });

  monthSelectContainer.appendChild(monthLabel);
  monthSelectContainer.appendChild(monthSelect);

  // Left: title + month selection; Right: close button
  const headerLeftContainer = document.createElement('div');
  headerLeftContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 16px;
  `;
  headerLeftContainer.appendChild(title);
  headerLeftContainer.appendChild(monthSelectContainer);

  headerTopRow.appendChild(headerLeftContainer);
  headerTopRow.appendChild(closeBtn);
  headerContainer.appendChild(headerTopRow);

  // Filter and sort container (merged into header, sticky)
  const filterContainer = document.createElement('div');
  filterContainer.style.cssText = `
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    padding: 12px 0 0 0;
    flex-wrap: wrap;
    width: 100%;
  `;
  
  // Add filters to headerContainer
  headerContainer.appendChild(filterContainer);

  // Left: filter type selection
  const filterTypeContainer = document.createElement('div');
  filterTypeContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 8px;
    flex-shrink: 0;
  `;

  const filterTypeLabel = document.createElement('label');
  filterTypeLabel.textContent = '篩選方式：';
  filterTypeLabel.style.cssText = 'font-size: 14px; color: #666; white-space: nowrap;';

  const filterTypeSelect = document.createElement('select');
  filterTypeSelect.id = 'history-filter-type-select';
  filterTypeSelect.style.cssText = `
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 14px;
    cursor: pointer;
    min-width: 100px;
  `;

  const filterTypeOptions = [
    { value: 'none', text: '無篩選' },
    { value: 'date', text: '日期' },
    { value: 'category', text: '消費類別' }
  ];

  filterTypeOptions.forEach(opt => {
    const option = document.createElement('option');
    option.value = opt.value;
    option.textContent = opt.text;
    filterTypeSelect.appendChild(option);
  });

  filterTypeContainer.appendChild(filterTypeLabel);
  filterTypeContainer.appendChild(filterTypeSelect);

  // 右邊：篩選值選擇（根據篩選方式動態更新）
  const filterValueContainer = document.createElement('div');
  filterValueContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 8px;
    flex: 1;
    justify-content: flex-end;
  `;

  const filterValueLabel = document.createElement('label');
  filterValueLabel.textContent = '篩選值：';
  filterValueLabel.style.cssText = 'font-size: 14px; color: #666; white-space: nowrap;';

  // 日期輸入框（用於日期篩選）
  const filterDateInput = document.createElement('input');
  filterDateInput.type = 'date';
  filterDateInput.id = 'history-filter-date-input';
  // 設置最小和最大日期來限制年份為四位數（1000-9999年）
  filterDateInput.min = '1000-01-01';
  filterDateInput.max = '9999-12-31';
  filterDateInput.style.cssText = `
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 14px;
    cursor: pointer;
    min-width: 150px;
  `;
  // 添加事件監聽器來驗證年份為四位數
  filterDateInput.addEventListener('input', (e) => {
    const value = e.target.value;
    if (value) {
      const year = parseInt(value.split('-')[0]);
      // 如果年份不是四位數，清空輸入或保持有效值
      if (isNaN(year) || year < 1000 || year > 9999) {
        // 只保留有效部分或清空
        const match = value.match(/^(\d{4})-\d{2}-\d{2}/);
        if (match && match[1].length === 4) {
          // 年份是四位數，保持原值
          return;
        } else {
          // 年份不是四位數，嘗試修正或清空
          const parts = value.split('-');
          if (parts[0].length > 4) {
            // 年份超過四位數，截取前四位
            parts[0] = parts[0].substring(0, 4);
            const correctedValue = parts.join('-');
            if (/^\d{4}-\d{2}-\d{2}/.test(correctedValue)) {
              e.target.value = correctedValue;
              filterValue = correctedValue;
              updateHistoryList(list, currentFilterRecords);
            }
          }
        }
      }
    }
  });

  // 類別下拉選單（用於類別篩選）
  const filterValueSelect = document.createElement('select');
  filterValueSelect.id = 'history-filter-value-select';
  filterValueSelect.style.cssText = `
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 14px;
    cursor: pointer;
    flex: 1;
    max-width: 300px;
  `;

  // 更新篩選值選項的函數（使用當前月份的記錄）
  const updateFilterValueOptions = (currentRecords) => {
    // 如果提供了記錄，更新當前記錄引用
    if (currentRecords) {
      currentFilterRecords = currentRecords;
    }
    
    filterValueSelect.innerHTML = '';
    filterValue = '';
    filterDateInput.value = '';

    if (filterType === 'date') {
      // 日期篩選：顯示日期輸入框
      filterDateInput.style.display = 'block';
      filterValueSelect.style.display = 'none';
    } else if (filterType === 'category') {
      // 消費類別篩選：顯示下拉選單，使用 EXPENSE_CATEGORY_OPTIONS
      filterDateInput.style.display = 'none';
      filterValueSelect.style.display = 'block';
      EXPENSE_CATEGORY_OPTIONS.forEach(cat => {
        const option = document.createElement('option');
        option.value = cat.value;
        option.textContent = cat.text;
        filterValueSelect.appendChild(option);
      });
    } else {
      // 無篩選：隱藏所有篩選值輸入
      filterDateInput.style.display = 'none';
      filterValueSelect.style.display = 'none';
      filterValueContainer.style.display = 'none';
      return;
    }

    // 顯示篩選值容器
    filterValueContainer.style.display = 'flex';
  };

  filterValueContainer.appendChild(filterValueLabel);
  filterValueContainer.appendChild(filterDateInput);
  filterValueContainer.appendChild(filterValueSelect);

  // 排序方式容器
  const sortContainer = document.createElement('div');
  sortContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 8px;
    flex-shrink: 0;
  `;

  const sortLabel = document.createElement('label');
  sortLabel.textContent = '排序方式：';
  sortLabel.style.cssText = 'font-size: 14px; color: #666; white-space: nowrap;';

  const sortSelect = document.createElement('select');
  sortSelect.id = 'history-sort-select';
  sortSelect.style.cssText = `
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 14px;
    cursor: pointer;
    min-width: 100px;
  `;

  const sortOptions = [
    { value: 'none', text: '無排序' },
    { value: 'date_desc', text: '日期：由大到小' },
    { value: 'date_asc', text: '日期：由小到大' },
    { value: 'category', text: '依類別' }
  ];

  sortOptions.forEach(opt => {
    const option = document.createElement('option');
    option.value = opt.value;
    option.textContent = opt.text;
    sortSelect.appendChild(option);
  });

  sortContainer.appendChild(sortLabel);
  sortContainer.appendChild(sortSelect);

  filterContainer.appendChild(filterTypeContainer);
  filterContainer.appendChild(filterValueContainer);
  filterContainer.appendChild(sortContainer);

  // 篩選方式變更事件
  filterTypeSelect.addEventListener('change', () => {
    filterType = filterTypeSelect.value;
    updateFilterValueOptions(); // 使用當前記錄
    // 重新應用篩選並更新列表（使用當前記錄）
    updateHistoryList(list, currentFilterRecords);
  });

  // 日期篩選值變更事件
  filterDateInput.addEventListener('change', () => {
    filterValue = filterDateInput.value;
    // 重新應用篩選並更新列表（使用當前記錄）
    updateHistoryList(list, currentFilterRecords);
  });

  // 類別篩選值變更事件
  filterValueSelect.addEventListener('change', () => {
    filterValue = filterValueSelect.value;
    // 重新應用篩選並更新列表（使用當前記錄）
    updateHistoryList(list, currentFilterRecords);
  });

  // 排序方式變更事件
  sortSelect.addEventListener('change', () => {
    sortType = sortSelect.value;
    // 重新應用篩選和排序並更新列表（使用當前記錄）
    updateHistoryList(list, currentFilterRecords);
  });

  // 初始狀態：隱藏篩選值容器
  updateFilterValueOptions();

  // 記錄列表
  const list = document.createElement('div');
  list.className = 'history-list';

  // 初始化記錄列表顯示
  updateHistoryList(list, records);

  // 只添加一次：headerContainer 已經包含篩選器
  content.appendChild(headerContainer);
  content.appendChild(list);
  modal.appendChild(content);

  // 點擊背景關閉
  modal.onclick = (e) => {
    if (e.target === modal) {
      modal.remove();
      isHistoryModalOpen = false;
      historyButton.disabled = false; // 重新啟用按鈕
    }
  };

  document.body.appendChild(modal);
}

// 顯示編輯彈出視窗
function showEditModal(record) {
  // 先從暫存區載入該月份的完整記錄列表
  const allRecordsForMonth = loadHistoryListFromCache(currentSheetIndex);
  // 根據記錄的唯一標識（時間和項目）找到對應的完整記錄
  const recordTime = record.row[0] || '';
  const recordItem = record.row[1] || '';
  const fullRecord = allRecordsForMonth.find(r => {
    const rTime = r.row[0] || '';
    const rItem = r.row[1] || '';
    return rTime === recordTime && rItem === recordItem;
  });

  if (!fullRecord) {
    alert('無法載入完整的記錄數據，請重新選擇');
    return;
  }
  const modal = document.createElement('div');
  modal.className = 'edit-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 10001;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;

  const content = document.createElement('div');
  content.className = 'edit-modal-content';
  content.style.cssText = `
    background-color: #fff;
    border-radius: 12px;
    padding: 30px;
    max-width: 800px;
    width: 100%;
    max-height: 90vh;
    overflow-y: auto;
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  // 關閉按鈕（右上角）
  const closeBtn = document.createElement('button');
  closeBtn.innerHTML = '✕';
  closeBtn.style.cssText = `
    position: absolute;
    top: 10px;
    right: 10px;
    width: 30px;
    height: 30px;
    border: none;
    background: transparent;
    font-size: 24px;
    cursor: pointer;
    color: #666;
    line-height: 1;
    z-index: 12;
  `;
  // 關閉按鈕：只關閉編輯視窗，回到歷史記錄列表
  closeBtn.onclick = () => {
    modal.remove();
    // 不關閉歷史記錄列表，讓用戶可以繼續選擇其他記錄
  };

  // 標題
  const title = document.createElement('h2');
  title.textContent = '編輯記錄';
  title.style.cssText = `
    margin: 0 0 20px 0;
    font-size: 20px;
    font-weight: 600;
  `;

  // 創建表單容器（複製表單結構）
  const formContainer = document.createElement('div');
  formContainer.className = 'edit-form-container';

  // 重新創建表單元素（避免ID衝突）
  // 日期欄位（可編輯）
  const dateRow = createInputRow('日期：', 'edit-date-input', 'date');
  // 為編輯表單的日期輸入框添加年份限制
  const editDateInputElement = dateRow.querySelector('#edit-date-input');
  if (editDateInputElement) {
    editDateInputElement.min = '1000-01-01';
    editDateInputElement.max = '9999-12-31';
    editDateInputElement.addEventListener('input', (e) => {
      const value = e.target.value;
      if (value) {
        const year = parseInt(value.split('-')[0]);
        if (!isNaN(year) && (year < 1000 || year > 9999)) {
          const parts = value.split('-');
          if (parts[0].length > 4) {
            parts[0] = parts[0].substring(0, 4);
            const correctedValue = parts.join('-');
            if (/^\d{4}-\d{2}-\d{2}/.test(correctedValue)) {
              e.target.value = correctedValue;
            }
          }
        }
      }
    });
  }
  const itemRow = createInputRow('項目：', 'edit-item-input');
  const expenseCategoryRow = createSelectRow('類別：', 'edit-expense-category-select', EXPENSE_CATEGORY_OPTIONS);

  const paymentMethodRow = createSelectRow('支付方式：', 'edit-payment-method-select', PAYMENT_METHOD_OPTIONS);

  const creditCardPaymentRow = createSelectRow('信用卡支付方式：', 'edit-credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);

  const monthPaymentRow = createSelectRow('本月／次月支付：', 'edit-month-payment-select', MONTH_PAYMENT_OPTIONS);

  const paymentPlatformRow = createSelectRow('支付平台：', 'edit-payment-platform-select', PAYMENT_PLATFORM_OPTIONS);

  const actualCostRow = createInputRow('實際消費金額：', 'edit-actual-cost-input', 'number');
  const recordCostRow = createInputRow('列帳消費金額：', 'edit-record-cost-input', 'number');

  // 備註使用 textarea
  const noteRow = document.createElement('div');
  noteRow.className = 'input-row';
  const noteLabel = document.createElement('label');
  noteLabel.textContent = '備註：';
  noteLabel.htmlFor = 'edit-note-input';
  const noteInput = document.createElement('textarea');
  noteInput.id = 'edit-note-input';
  noteInput.name = 'edit-note-input';
  noteInput.rows = 3;
  noteRow.appendChild(noteLabel);
  noteRow.appendChild(noteInput);

  // 設置條件顯示的初始狀態（預設隱藏）
  creditCardPaymentRow.style.display = 'none';
  monthPaymentRow.style.display = 'none';
  paymentPlatformRow.style.display = 'none';

  // 添加支付方式變更事件處理，控制條件欄位的顯示/隱藏
  const paymentMethodSelect = paymentMethodRow.querySelector('#edit-payment-method-select');
  if (paymentMethodSelect) {
    paymentMethodSelect.addEventListener('change', () => {
      const paymentMethod = paymentMethodSelect.value;
      if (isCreditCardPayment(paymentMethod)) {
        creditCardPaymentRow.style.display = 'flex';
        monthPaymentRow.style.display = 'flex';
        paymentPlatformRow.style.display = 'none';
      } else if (isStoredValuePayment(paymentMethod)) {
        creditCardPaymentRow.style.display = 'none';
        monthPaymentRow.style.display = 'none';
        paymentPlatformRow.style.display = 'flex';
  } else {
        creditCardPaymentRow.style.display = 'none';
        monthPaymentRow.style.display = 'none';
        paymentPlatformRow.style.display = 'none';
      }
    });
  }

  formContainer.appendChild(dateRow);
  formContainer.appendChild(itemRow);
  formContainer.appendChild(expenseCategoryRow);
  formContainer.appendChild(paymentMethodRow);
  formContainer.appendChild(creditCardPaymentRow);
  formContainer.appendChild(monthPaymentRow);
  formContainer.appendChild(paymentPlatformRow);
  formContainer.appendChild(actualCostRow);
  formContainer.appendChild(recordCostRow);
  formContainer.appendChild(noteRow);

  // 先將表單添加到 modal，確保元素在 DOM 中
  content.appendChild(closeBtn);
  content.appendChild(title);
  content.appendChild(formContainer);
  modal.appendChild(content);
  document.body.appendChild(modal);

  // 等待 DOM 更新後再填充數據
  setTimeout(() => {
    // 使用完整記錄數據填充表單
    const row = fullRecord.row;

  const dateInput = document.getElementById('edit-date-input');
  const itemInput = document.getElementById('edit-item-input');
  const expenseCategorySelect = document.getElementById('edit-expense-category-select');
  // paymentMethodSelect 已經在第 1858 行聲明，不需要重複聲明
  const creditCardPaymentSelect = document.getElementById('edit-credit-card-payment-select');
  const monthPaymentSelect = document.getElementById('edit-month-payment-select');
  const paymentPlatformSelect = document.getElementById('edit-payment-platform-select');
  const actualCostInput = document.getElementById('edit-actual-cost-input');
  const recordCostInput = document.getElementById('edit-record-cost-input');
  const noteInput = document.getElementById('edit-note-input'); // 重新獲取，確保元素已存在

  // 填充日期（確保可以編輯）
  if (dateInput) {
    // 日期格式轉換：將各種日期格式轉換為 "YYYY-MM-DD"
    // row[0] 可能是 Date 對象、字符串或其他格式，需要先轉換為字符串
    let dateValue = row[0] || '';
    if (dateValue instanceof Date) {
      // 如果是 Date 對象，轉換為字符串
      dateValue = dateValue.toISOString();
    } else {
      // 確保是字符串
      dateValue = String(dateValue);
    }
    
    let formattedDate = '';
    if (dateValue && dateValue.trim() !== '') {
      // 處理 ISO 格式（包含時間部分，如 "2026-01-16T16:00:00.000Z"）
      if (dateValue.includes('T') || dateValue.includes('Z')) {
        const dateObj = new Date(dateValue);
        if (!isNaN(dateObj.getTime())) {
          const year = dateObj.getFullYear();
          const month = String(dateObj.getMonth() + 1).padStart(2, '0');
          const day = String(dateObj.getDate()).padStart(2, '0');
          formattedDate = `${year}-${month}-${day}`;
        }
      }
      // 處理 "YYYY/MM/DD" 格式
      else if (dateValue.includes('/')) {
        const parts = dateValue.split('/');
        if (parts.length >= 3) {
          const year = parts[0].trim();
          const month = parts[1].trim().padStart(2, '0');
          const day = parts[2].trim().split(' ')[0].split('T')[0].padStart(2, '0');
          formattedDate = `${year}-${month}-${day}`;
        }
      }
      // 處理 "YYYY-MM-DD" 格式（可能包含時間部分）
      else if (dateValue.includes('-')) {
        // 提取日期部分（如果包含時間，取前面的日期部分）
        const datePart = dateValue.split(' ')[0].split('T')[0];
        const parts = datePart.split('-');
        if (parts.length >= 3) {
          const year = parts[0].trim();
          const month = parts[1].trim().padStart(2, '0');
          const day = parts[2].trim().padStart(2, '0');
          formattedDate = `${year}-${month}-${day}`;
        }
      } else {
        // 嘗試解析其他格式
        const dateObj = new Date(dateValue);
        if (!isNaN(dateObj.getTime())) {
          const year = dateObj.getFullYear();
          const month = String(dateObj.getMonth() + 1).padStart(2, '0');
          const day = String(dateObj.getDate()).padStart(2, '0');
          formattedDate = `${year}-${month}-${day}`;
        }
      }
    }
    
    console.log('[Expense] Date formatting:', {
      original: row[0],
      dateValue: dateValue,
      formattedDate: formattedDate
    });
    
    dateInput.value = formattedDate;
    dateInput.readOnly = false; // 確保可以編輯
    dateInput.disabled = false; // 確保可以編輯
  }

  // 填充項目（確保可以編輯）
  if (itemInput) {
    itemInput.value = row[1] || '';
    itemInput.readOnly = false; // 確保可以編輯
    itemInput.disabled = false; // 確保可以編輯
  } else {
  }

  // 填充類別
  if (expenseCategorySelect) {
    expenseCategorySelect.value = row[2] || '';
    const selectContainer = expenseCategorySelect.parentElement;
    if (selectContainer) {
      const selectDisplay = selectContainer.querySelector('.select-display');
      if (selectDisplay) {
        const selectText = selectDisplay.querySelector('.select-text');
        if (selectText) {
          const selectedOption = expenseCategorySelect.options[expenseCategorySelect.selectedIndex];
          selectText.textContent = selectedOption ? selectedOption.textContent : row[2] || '';
        }
      }
    }
  }

  // 填充支付方式（先設置值，觸發 change 事件以顯示/隱藏條件欄位）
  if (paymentMethodSelect) {
    paymentMethodSelect.value = row[3] || '';
    const selectContainer = paymentMethodSelect.parentElement;
    if (selectContainer) {
      const selectDisplay = selectContainer.querySelector('.select-display');
      if (selectDisplay) {
        const selectText = selectDisplay.querySelector('.select-text');
        if (selectText) {
          const selectedOption = paymentMethodSelect.options[paymentMethodSelect.selectedIndex];
          selectText.textContent = selectedOption ? selectedOption.textContent : row[3] || '';
        }
      }
    }
    // 手動設置條件欄位的顯示狀態
    const paymentMethod = row[3] || '';
    if (isCreditCardPayment(paymentMethod)) {
      creditCardPaymentRow.style.display = 'flex';
      monthPaymentRow.style.display = 'flex';
      paymentPlatformRow.style.display = 'none';
    } else if (isStoredValuePayment(paymentMethod)) {
      creditCardPaymentRow.style.display = 'none';
      monthPaymentRow.style.display = 'none';
      paymentPlatformRow.style.display = 'flex';
    } else {
      creditCardPaymentRow.style.display = 'none';
      monthPaymentRow.style.display = 'none';
      paymentPlatformRow.style.display = 'none';
    }

    // 等待 DOM 更新後再填充條件欄位
    setTimeout(() => {
      // 填充信用卡支付方式（如果支付方式是信用卡類型）
      if (isCreditCardPayment(paymentMethod) && creditCardPaymentSelect) {
        creditCardPaymentSelect.value = row[4] || '';
        const selectContainer = creditCardPaymentSelect.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
              const selectedOption = creditCardPaymentSelect.options[creditCardPaymentSelect.selectedIndex];
              selectText.textContent = selectedOption ? selectedOption.textContent : row[4] || '';
            }
          }
        }
      }

      // 填充本月/次月支付（如果支付方式是信用卡類型）
      if (isCreditCardPayment(paymentMethod) && monthPaymentSelect) {
        monthPaymentSelect.value = row[5] || '';
        const selectContainer = monthPaymentSelect.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
              const selectedOption = monthPaymentSelect.options[monthPaymentSelect.selectedIndex];
              selectText.textContent = selectedOption ? selectedOption.textContent : row[5] || '';
            }
          }
        }
      }

      // 填充支付平台（如果支付方式是存款或儲值類型）
      if (isStoredValuePayment(paymentMethod) && paymentPlatformSelect) {
        paymentPlatformSelect.value = row[7] || '';
        const selectContainer = paymentPlatformSelect.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
              const selectedOption = paymentPlatformSelect.options[paymentPlatformSelect.selectedIndex];
              selectText.textContent = selectedOption ? selectedOption.textContent : row[7] || '';
            }
          }
        }
      }
    }, 50);
  }

  // 填充金額和備註（在 setTimeout 內部，確保元素已存在）
  if (actualCostInput) {
    actualCostInput.value = row[6] || '';
  } else {
  }
  if (recordCostInput) {
    recordCostInput.value = row[8] || '';
  } else {
  }
  if (noteInput) {
    noteInput.value = row[9] || '';
  } else {
  }
  }, 100); // 等待 DOM 完全渲染後再填充

  // 儲存按鈕
  const editModalSaveButton = document.createElement('button');
  editModalSaveButton.textContent = '儲存';
  editModalSaveButton.className = 'save-button';
  editModalSaveButton.style.cssText = `
    margin-top: 20px;
    width: 100%;
  `;

  let isSaving = false;
  editModalSaveButton.onclick = async () => {
    if (isSaving) return;
    isSaving = true;

    // 鎖定整個頁面，等待後端回傳
    showSpinner();
    editModalSaveButton.textContent = '儲存中...';
    editModalSaveButton.disabled = true;

    // 禁用所有輸入和按鈕
    const editDateInput = document.getElementById('edit-date-input');
    const editItemInput = document.getElementById('edit-item-input');
    const editExpenseCategorySelect = document.getElementById('edit-expense-category-select');
    const editPaymentMethodSelect = document.getElementById('edit-payment-method-select');
    const editCreditCardPaymentSelect = document.getElementById('edit-credit-card-payment-select');
    const editMonthPaymentSelect = document.getElementById('edit-month-payment-select');
    const editPaymentPlatformSelect = document.getElementById('edit-payment-platform-select');
    const editActualCostInput = document.getElementById('edit-actual-cost-input');
    const editRecordCostInput = document.getElementById('edit-record-cost-input');
    const editNoteInput = document.getElementById('edit-note-input');
    if (editDateInput) editDateInput.disabled = true;
    if (editItemInput) editItemInput.disabled = true;
    if (editExpenseCategorySelect) editExpenseCategorySelect.disabled = true;
    if (editPaymentMethodSelect) editPaymentMethodSelect.disabled = true;
    if (editCreditCardPaymentSelect) editCreditCardPaymentSelect.disabled = true;
    if (editMonthPaymentSelect) editMonthPaymentSelect.disabled = true;
    if (editPaymentPlatformSelect) editPaymentPlatformSelect.disabled = true;
    if (editActualCostInput) editActualCostInput.disabled = true;
    if (editRecordCostInput) editRecordCostInput.disabled = true;
    if (editNoteInput) editNoteInput.disabled = true;

    // 禁用主頁面的所有輸入和按鈕
    const itemInput = document.getElementById('item-input');
    const expenseCategorySelect = document.getElementById('expense-category-select');
    const paymentMethodSelect = document.getElementById('payment-method-select');
    const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
    const monthPaymentSelect = document.getElementById('month-payment-select');
    const paymentPlatformSelect = document.getElementById('payment-platform-select');
    const actualCostInput = document.getElementById('actual-cost-input');
    const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.disabled = true;
    if (expenseCategorySelect) expenseCategorySelect.disabled = true;
    if (paymentMethodSelect) paymentMethodSelect.disabled = true;
    if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = true;
    if (monthPaymentSelect) monthPaymentSelect.disabled = true;
    if (paymentPlatformSelect) paymentPlatformSelect.disabled = true;
    if (actualCostInput) actualCostInput.disabled = true;
    if (recordCostInput) recordCostInput.disabled = true;
    if (noteInput) noteInput.disabled = true;
    if (mainSaveButton) mainSaveButton.disabled = true;
    if (historyButton) historyButton.disabled = true;

    try {
      await saveDataForEdit(fullRecord);
      // 等待後端回傳後才顯示成功訊息
      alert('修改成功！');
      modal.remove();
      // 關閉編輯視窗後，重新載入歷史記錄列表以顯示最新數據
      // 找到歷史記錄列表的 modal
      const historyModal = document.querySelector('.history-modal');
      if (historyModal) {
        // 重新載入當前月份的記錄列表（從暫存區重新載入，因為已經更新）
        const newRecords = loadHistoryListFromCache(currentSheetIndex);
        const listElement = historyModal.querySelector('.history-list');
        if (listElement) {
          // 使用 updateHistoryList 函數更新列表
          listElement.innerHTML = '';
          const displayRecords = newRecords;
          if (displayRecords.length === 0) {
            const emptyMsg = document.createElement('div');
            emptyMsg.textContent = '尚無歷史紀錄';
            emptyMsg.style.cssText = 'text-align: center; padding: 40px; color: #999;';
            listElement.appendChild(emptyMsg);
          } else {
            displayRecords.forEach(record => {
              listElement.appendChild(createHistoryItem(record, listElement));
            });
          }
        }
      }
    } catch (error) {
      alert('儲存失敗: ' + error.message);
    } finally {
      // 恢復所有按鈕和輸入
      hideSpinner();
      isSaving = false;
      editModalSaveButton.textContent = '儲存';
      editModalSaveButton.disabled = false;

      // 恢復編輯 modal 的輸入
      if (editDateInput) editDateInput.disabled = false;
      if (editItemInput) editItemInput.disabled = false;
      if (editExpenseCategorySelect) editExpenseCategorySelect.disabled = false;
      if (editPaymentMethodSelect) editPaymentMethodSelect.disabled = false;
      if (editCreditCardPaymentSelect) editCreditCardPaymentSelect.disabled = false;
      if (editMonthPaymentSelect) editMonthPaymentSelect.disabled = false;
      if (editPaymentPlatformSelect) editPaymentPlatformSelect.disabled = false;
      if (editActualCostInput) editActualCostInput.disabled = false;
      if (editRecordCostInput) editRecordCostInput.disabled = false;
      if (editNoteInput) editNoteInput.disabled = false;

      // 恢復主頁面的輸入和按鈕
      if (itemInput) itemInput.disabled = false;
      if (expenseCategorySelect) expenseCategorySelect.disabled = false;
      if (paymentMethodSelect) paymentMethodSelect.disabled = false;
      if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
      if (monthPaymentSelect) monthPaymentSelect.disabled = false;
      if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
      if (actualCostInput) actualCostInput.disabled = false;
      if (recordCostInput) recordCostInput.disabled = false;
      if (noteInput) noteInput.disabled = false;
      if (mainSaveButton) mainSaveButton.disabled = false;
      if (historyButton) historyButton.disabled = false;
    }
  };

  content.appendChild(closeBtn);
  content.appendChild(title);
  content.appendChild(formContainer);
  content.appendChild(editModalSaveButton);
  modal.appendChild(content);

  // 點擊背景關閉編輯視窗，返回歷史記錄列表
  modal.onclick = (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  };

  document.body.appendChild(modal);
}

// 刪除記錄的函數
async function deleteRecord(record) {
  // 優先使用預先標註好的 sheetRowIndex（在 loadHistoryListFromCache 中設定）
  const row = record.sheetRowIndex;

  if (!row) {
    throw new Error('無法找到要刪除的記錄位置（缺少列號資訊）');
  }

  const result = await callAPI({
    name: "Delete Data",
    sheet: currentSheetIndex,
    // 後端統一使用 updateRow 作為列號參數（新增/修改/刪除共用）
    updateRow: row
  });

  // 更新總計顯示（使用最新資料重新計算）
  updateTotalDisplay();

  // 更新暫存區和歷史記錄（使用回傳的 data）
  if (result.data) {
    processDataFromResponse(result.data, false, currentSheetIndex);
    allMonthsData[currentSheetIndex] = { data: result.data, total: result.total };

    // 保存到 IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});

    refreshHistoryList();
  }
}

// 編輯模式的儲存函數
async function saveDataForEdit(record) {
  // Get payment method first to check if it's credit card
  const editPaymentMethodSelect = document.getElementById('edit-payment-method-select');
  const paymentMethodValue = editPaymentMethodSelect ? editPaymentMethodSelect.value : '';
  
  // 如果是信用卡，確保正確取得月份支付選擇值
  let monthValue = '';
  if (isCreditCardPayment(paymentMethodValue)) {
      // 編輯模式：直接取得隱藏的 select 元素值
    const monthSelectElement = document.getElementById('edit-month-payment-select');
    if (monthSelectElement && monthSelectElement.tagName === 'SELECT') {
      monthValue = monthSelectElement.value || '';
      console.log('[Expense] Edit mode - Directly getting month payment value:', {
        element: monthSelectElement,
        value: monthValue,
        selectedIndex: monthSelectElement.selectedIndex,
        options: Array.from(monthSelectElement.options).map(o => ({ value: o.value, text: o.text, selected: o.selected }))
      });
    } else {
      console.warn('[Expense] Edit mode - edit-month-payment-select element not found or not a SELECT:', monthSelectElement);
    }
  }
  
  const formData = getFormData('edit');
  
  // 如果是信用卡支付，用直接取得的值覆蓋 monthIndex
  if (isCreditCardPayment(formData.spendWay) && monthValue) {
    formData.monthIndex = monthValue;
    console.log('[Expense] Edit mode - Override monthIndex:', {
      spendWay: formData.spendWay,
      monthValue: monthValue,
      formDataMonthIndex: formData.monthIndex
    });
  }

  const allRecordsForMonth = loadHistoryListFromCache(currentSheetIndex);
  const recordTime = formatRecordDateTime(record.row[0] || '');
  const recordItem = (record.row[1] || '').trim();
  const recordCategory = (record.row[2] || '').trim();
  const originalRecordCost = (record.row[6] || '').toString().trim();

  const recordIndex = allRecordsForMonth.findIndex(r => {
    const rTime = formatRecordDateTime(r.row[0] || '');
    const rItem = (r.row[1] || '').trim();
    const rCategory = (r.row[2] || '').trim();
    const rCost = (r.row[6] || '').toString().trim();
    return rTime === recordTime && rItem === recordItem && rCategory === recordCategory && rCost === originalRecordCost;
  });

  if (recordIndex === -1) {
    throw new Error('無法找到對應的記錄位置');
  }

  // 將 monthIndex 轉換為 month 供後端 API 使用（後端期望 'month' 而非 'monthIndex'）
  const apiData = {
    name: "Upsert Data",
    sheet: currentSheetIndex,
    date: formData.date,
    item: formData.item,
    category: formData.category,
    spendWay: formData.spendWay,
    creditCard: formData.creditCard,
    month: formData.monthIndex || '', // 後端期望 'month' 但前端使用 'monthIndex'
    actualCost: formData.actualCost,
    payment: formData.payment,
    recordCost: formData.recordCost,
    note: formData.note,
    // sheet 第 1 列是標題，第 2 列才是第一筆資料，所以要 +2
    updateRow: recordIndex + 2
  };
  
    // 編輯模式的調試日誌
  if (isCreditCardPayment(formData.spendWay)) {
    console.log('[Expense] Edit mode - Saving credit card expense:', {
      spendWay: formData.spendWay,
      creditCard: formData.creditCard,
      monthIndex: formData.monthIndex,
      month: apiData.month,
      monthValue: monthValue,
      apiData: apiData
    });
  }

  const result = await callAPI(apiData);

  // 更新總計顯示（使用最新資料重新計算）
  updateTotalDisplay();

  // 更新暫存區和歷史記錄（使用回傳的 data）
  if (result.data) {
    processDataFromResponse(result.data, false, currentSheetIndex);
    allMonthsData[currentSheetIndex] = { data: result.data, total: result.total };

    // 保存到 IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});

    refreshHistoryList();
  }
}

// 切換收入/支出類型
function switchType(targetType) {
    // 確保 categorySelect 元素存在
    const categorySelectElement = document.getElementById('category-select');
    if (!categorySelectElement) {
      return; // 如果元素不存在，直接返回
    }

    const currentType = categorySelectElement.value || '支出';

    // 如果目標類型與當前類型不同，則切換
    if (currentType !== targetType) {
      // 先更新全局變數，這樣 updateDivVisibility 才能讀取到正確的值
      if (typeof categorySelect !== 'undefined') {
        categorySelect.value = targetType;
      }

      // 更新 DOM 元素
      categorySelectElement.value = targetType;

      // 更新顯示文字
      const selectContainer = categorySelectElement.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('div');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('div');
          if (selectText) {
            selectText.textContent = targetType;
          }
        }
      }

      // 先更新 UI 元素顯示（特別是支出/收入的欄位切換）
      if (typeof updateDivVisibility === 'function') {
        updateDivVisibility();
      }

      // 然後過濾記錄並更新 UI（filterRecordsByType 會自動處理相同編號的查找）
      // 使用 setTimeout 確保 updateDivVisibility 完成後再執行
      setTimeout(() => {
      // updateDivVisibility 會重新創建 expense-category-select 元素，需要重新獲取並更新
      const newCategorySelectElement = document.getElementById('expense-category-select');
        if (newCategorySelectElement) {
          // 更新新創建的元素的值（如果目標類型是支出，設置第一個選項；如果是收入，不需要設置）
          if (targetType === '支出' && newCategorySelectElement.options.length > 0) {
            newCategorySelectElement.value = newCategorySelectElement.options[0].value;
            // 同步更新自訂下拉顯示文字
            const newSelectContainer = newCategorySelectElement.parentElement;
            if (newSelectContainer) {
              const newSelectDisplay = newSelectContainer.querySelector('div');
              if (newSelectDisplay) {
                const newSelectText = newSelectDisplay.querySelector('div');
                if (newSelectText) {
                  newSelectText.textContent = newCategorySelectElement.options[0].textContent;
                }
              }
            }
          }

          // 更新全局變數的 value 屬性（雖然元素已更換，但我們可以通過更新屬性來保持一致性）
          // 實際上，由於 categorySelect 是 const，我們需要確保後續代碼使用 getElementById 獲取最新元素
          // 但為了兼容性，我們也更新全局變數的 value（如果元素還存在的話）
          if (typeof categorySelect !== 'undefined' && categorySelect.parentNode) {
            categorySelect.value = targetType;
          }
        }

        if (typeof filterRecordsByType === 'function') {
          filterRecordsByType(targetType);
        }

        // 更新相關 UI 元素
        if (typeof updateDeleteButton === 'function') {
          updateDeleteButton();
        }
      }, 200);
  }
}

// 鍵盤事件已簡化，移除導航功能

// 觸摸滑動事件已移除（純新增模式不需要）

// 移除類別選擇（支出/收入），改為新的表單欄位結構

const itemContainer = document.createElement('div');
itemContainer.className = 'item-container';

// 0. 日期（所有類型都要填寫）
const dateRow = createInputRow('日期：', 'date-input', 'date');
const dateInput = dateRow.querySelector('#date-input');
if (dateInput) {
  // 設置年份限制為四位數（1000-9999年）
  dateInput.min = '1000-01-01';
  dateInput.max = '9999-12-31';
  // 添加事件監聽器來驗證年份為四位數
  dateInput.addEventListener('input', (e) => {
    const value = e.target.value;
    if (value) {
      const year = parseInt(value.split('-')[0]);
      if (!isNaN(year) && (year < 1000 || year > 9999)) {
        const parts = value.split('-');
        if (parts[0].length > 4) {
          parts[0] = parts[0].substring(0, 4);
          const correctedValue = parts.join('-');
          if (/^\d{4}-\d{2}-\d{2}/.test(correctedValue)) {
            e.target.value = correctedValue;
          }
        }
      }
    }
  });
  // 預設帶入今天（YYYY-MM-DD），可以手動修改
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  dateInput.value = `${year}-${month}-${day}`;

  // 當日期變更時，切換到對應月份並更新 summary
  dateInput.addEventListener('change', async () => {
    const selectedDate = dateInput.value; // 格式：YYYY-MM-DD
    if (!selectedDate) return;

    const [yearStr, monthStr] = selectedDate.split('-');
    const targetMonthStr = `${yearStr}${monthStr}`; // 例如："202501"

    // 找到對應月份的 sheetIndex
    const targetIndex = sheetNames.findIndex(name => name === targetMonthStr);
    if (targetIndex !== -1) {
      const newSheetIndex = targetIndex + 2; // sheetIndex = index + 2（因為前兩個是無效項目）
      if (newSheetIndex !== currentSheetIndex) {
        currentSheetIndex = newSheetIndex;
        // 載入該月份資料並更新 summary
        if (!allMonthsData[currentSheetIndex]) {
          try {
            const monthData = await loadMonthData(currentSheetIndex);
            allMonthsData[currentSheetIndex] = monthData;
            setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});
          } catch (e) {
          }
        }
        updateTotalDisplay();
        // 如果函數存在，更新月份指示器
        if (typeof updateMonthIndicator === 'function') {
          updateMonthIndicator();
        }
      }
    }
  });
}

// 1. 項目
const itemRow = createInputRow('項目：', 'item-input', 'text');
const itemInput = itemRow.querySelector('#item-input');
itemInput.placeholder = '輸入項目名稱...';

// 2. 類別（消費類別）- 使用統一常數
const categoryRow = createSelectRow('類別：', 'expense-category-select', EXPENSE_CATEGORY_OPTIONS);
const expenseCategorySelect = categoryRow.querySelector('#expense-category-select');

// 類別變更時即時更新上方「預算 / 支出 / 餘額」
if (expenseCategorySelect) {
  expenseCategorySelect.addEventListener('change', () => {
    updateTotalDisplay();
  });
}

// 3. 支付方式 - 使用統一常數
const paymentMethodRow = createSelectRow('支付方式：', 'payment-method-select', PAYMENT_METHOD_OPTIONS);
const paymentMethodSelect = paymentMethodRow.querySelector('#payment-method-select');

// 4. 信用卡支付方式（條件顯示：支付方式是信用卡）- 使用統一常數
const creditCardPaymentRow = createSelectRow('信用卡支付方式：', 'credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);
creditCardPaymentRow.style.display = 'none'; // 預設隱藏
const creditCardPaymentSelect = creditCardPaymentRow.querySelector('#credit-card-payment-select');

// 5. 本月／次月支付（條件顯示：支付方式是信用卡）- 使用統一常數
const monthPaymentRow = createSelectRow('本月／次月支付：', 'month-payment-select', MONTH_PAYMENT_OPTIONS);
monthPaymentRow.style.display = 'none'; // 預設隱藏
const monthPaymentSelect = monthPaymentRow.querySelector('#month-payment-select');

// 6. 支付平台（條件顯示：支付方式是存款或儲值的支出）- 使用統一常數
const paymentPlatformRow = createSelectRow('支付平台：', 'payment-platform-select', PAYMENT_PLATFORM_OPTIONS);
paymentPlatformRow.style.display = 'none'; // 預設隱藏
const paymentPlatformSelect = paymentPlatformRow.querySelector('#payment-platform-select');

// 7. 實際消費金額
const actualCostRow = createInputRow('實際消費金額：', 'actual-cost-input', 'number');
const actualCostInput = actualCostRow.querySelector('#actual-cost-input');

// 8. 列帳消費金額
const recordCostRow = createInputRow('列帳消費金額：', 'record-cost-input', 'number');
const recordCostInput = recordCostRow.querySelector('#record-cost-input');

// 輸入處理器的防抖輔助函數
let updateTotalDebounceTimer = null;
const debouncedUpdateTotal = (immediate = false) => {
  if (updateTotalDebounceTimer) {
    clearTimeout(updateTotalDebounceTimer);
  }
  if (immediate) {
    updateTotalDisplay();
  } else {
    updateTotalDebounceTimer = setTimeout(() => {
      updateTotalDisplay();
    }, 150); // 150ms debounce
  }
};

// 實際金額 / 列帳金額 / 類別變更時即時更新上方「預算 / 支出 / 餘額」
if (actualCostInput) {
  actualCostInput.addEventListener('input', () => {
    debouncedUpdateTotal();
  });
}
if (recordCostInput) {
  recordCostInput.addEventListener('input', () => {
    debouncedUpdateTotal();
  });

  // 當列帳金額輸入框獲得焦點且總計區域不可見時，使總計區域固定
  const placeholder = document.createElement('div');
  placeholder.className = 'total-container-placeholder';
  placeholder.style.cssText = 'display: none; width: 100%;';

  let stickyActive = false;
  let inputFocused = false;

  // 動態取得實際的頁首高度
  const getHeaderHeight = () => {
    const header = document.querySelector('.site-header');
    return header ? header.offsetHeight : 56;
  };

  // 檢查總計區域是否在視口中可見
  const isSummaryVisible = () => {
    if (!placeholder.parentNode) return true; // 如果沒有佔位符，檢查原始位置
    const rect = placeholder.getBoundingClientRect();
    const vv = window.visualViewport;
    const viewportTop = vv ? vv.offsetTop : 0;
    const navHeight = getHeaderHeight();
    // 如果總計區域的頂部在導航欄下方，則總計區域可見
    return rect.top >= (viewportTop + navHeight);
  };

  // 根據滾動位置和可見性更新固定狀態
  const updateStickyState = () => {
    if (!inputFocused || !totalContainer) return;

    const shouldBeSticky = !isSummaryVisible();

    if (shouldBeSticky && !stickyActive) {
      // 設為固定
      placeholder.style.height = totalContainer.offsetHeight + 'px';
      placeholder.style.display = 'block';
      totalContainer.classList.add('sticky-active');
      stickyActive = true;
      updateStickyPosition();
    } else if (!shouldBeSticky && stickyActive) {
      // 移除固定
      totalContainer.classList.remove('sticky-active');
      totalContainer.style.top = '';
      placeholder.style.display = 'none';
      stickyActive = false;
    }
  };

  // 根據視覺視口更新固定位置（用於鍵盤）
  const updateStickyPosition = () => {
    if (!stickyActive || !totalContainer) return;

    const vv = window.visualViewport;
    const headerHeight = getHeaderHeight();
    if (vv) {
      // 檢查鍵盤是否可能打開（視口高度顯著減少）
      const keyboardOpen = vv.height < window.innerHeight * 0.75;

      if (keyboardOpen) {
        // 當鍵盤打開時，定位在視覺視口頂部
        totalContainer.style.top = vv.offsetTop + 'px';
      } else {
        // 正常情況：定位在導航欄下方
        totalContainer.style.top = headerHeight + 'px';
      }
    } else {
      totalContainer.style.top = headerHeight + 'px';
    }
  };

  // 監聽視覺視口變化（鍵盤顯示/隱藏）
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', updateStickyPosition);
    window.visualViewport.addEventListener('scroll', updateStickyPosition);
  }

  // 監聽滾動以檢查可見性
  window.addEventListener('scroll', updateStickyState, { passive: true });

  recordCostInput.addEventListener('focus', () => {
    if (totalContainer && totalContainer.parentNode) {
      inputFocused = true;
      // 如果佔位符尚未存在，在 totalContainer 之前插入佔位符
      if (!placeholder.parentNode) {
        totalContainer.parentNode.insertBefore(placeholder, totalContainer);
      }
      placeholder.style.height = totalContainer.offsetHeight + 'px';
      placeholder.style.display = 'block';
      // 立即檢查是否應該固定
      updateStickyState();
    }
  });

  recordCostInput.addEventListener('blur', () => {
    inputFocused = false;
    if (totalContainer) {
      totalContainer.classList.remove('sticky-active');
      totalContainer.style.top = '';
      placeholder.style.display = 'none';
      stickyActive = false;
    }
  });
}

// 9. 備註
const noteRow = document.createElement('div');
noteRow.className = 'input-row';
const noteLabel = document.createElement('label');
noteLabel.textContent = '備註：';
noteLabel.htmlFor = 'note-input';
const noteInput = document.createElement('textarea');
noteInput.id = 'note-input';
noteInput.name = 'note-input';
noteInput.rows = 3;
noteRow.appendChild(noteLabel);
noteRow.appendChild(noteInput);

// 條件顯示邏輯：根據支付方式顯示/隱藏相關欄位
const updatePaymentFieldsVisibility = () => {
  const paymentMethod = paymentMethodSelect.value;

  // 如果支付方式名稱含有「信用卡」等關鍵字，顯示信用卡支付方式和本月/次月支付
  if (isCreditCardPayment(paymentMethod)) {
    creditCardPaymentRow.style.display = 'flex';
    monthPaymentRow.style.display = 'flex';
    paymentPlatformRow.style.display = 'none';
  }
  // 如果支付方式名稱含有「存款」、「儲值」等關鍵字，顯示支付平台
  else if (isStoredValuePayment(paymentMethod)) {
    creditCardPaymentRow.style.display = 'none';
    monthPaymentRow.style.display = 'none';
    paymentPlatformRow.style.display = 'flex';
  }
  // 其他情況隱藏所有條件欄位
  else {
    creditCardPaymentRow.style.display = 'none';
    monthPaymentRow.style.display = 'none';
    paymentPlatformRow.style.display = 'none';
  }
};

// 監聽支付方式變更
paymentMethodSelect.addEventListener('change', updatePaymentFieldsVisibility);

// 將所有欄位添加到 itemContainer（日期放在項目名稱下面）
itemContainer.appendChild(itemRow);
itemContainer.appendChild(dateRow);
itemContainer.appendChild(categoryRow);
itemContainer.appendChild(paymentMethodRow);
itemContainer.appendChild(creditCardPaymentRow);
itemContainer.appendChild(monthPaymentRow);
itemContainer.appendChild(paymentPlatformRow);
itemContainer.appendChild(actualCostRow);
itemContainer.appendChild(recordCostRow);
itemContainer.appendChild(noteRow);

const mainSaveButton = document.createElement('button');
mainSaveButton.textContent = '儲存';
mainSaveButton.className = 'save-button';

const columnsContainer = document.createElement('div');
columnsContainer.className = 'columns-container';
columnsContainer.style.position = 'relative';

// 總計區域載入中覆蓋層
const summaryLoadingOverlay = document.createElement('div');
summaryLoadingOverlay.className = 'summary-loading-overlay';
summaryLoadingOverlay.innerHTML = '<span class="summary-spinner"></span>';
columnsContainer.appendChild(summaryLoadingOverlay);

// 總計區域中的月份指示器
const monthIndicator = document.createElement('div');
monthIndicator.className = 'summary-month-indicator';
monthIndicator.style.cssText = 'font-size: 14px; color: #666; margin-bottom: 8px; text-align: center; font-weight: 500;';

// 根據 currentSheetIndex 更新月份指示器的函數
const updateMonthIndicator = () => {
  if (sheetNames.length > 0 && currentSheetIndex >= 2) {
    const monthName = sheetNames[currentSheetIndex - 2]; // 例如："202501"
    if (monthName && monthName.length >= 6) {
      const year = monthName.substring(0, 4);
      const month = monthName.substring(4, 6);
      monthIndicator.textContent = `${year}年${parseInt(month)}月`;
    }
  } else {
    // 預設為當前月份
    const now = new Date();
    monthIndicator.textContent = `${now.getFullYear()}年${now.getMonth() + 1}月`;
  }
};
updateMonthIndicator();

const incomeColumn = document.createElement('div');
incomeColumn.className = 'income-column';

const incomeTitle = document.createElement('h3');
incomeTitle.className = 'income-title';
incomeTitle.textContent = '預算：';

const incomeAmount = document.createElement('div');
incomeAmount.className = 'income-amount';
incomeAmount.textContent = '---';

const expenseColumn = document.createElement('div');
expenseColumn.className = 'expense-column';

const expenseTitle = document.createElement('h3');
expenseTitle.className = 'expense-title';
expenseTitle.textContent = '支出：';

const expenseAmount = document.createElement('div');
expenseAmount.className = 'expense-amount';
expenseAmount.textContent = '---';

const totalColumn = document.createElement('div');
totalColumn.className = 'total-column';

const totalTitle = document.createElement('h3');
totalTitle.className = 'total-title';
totalTitle.textContent = '餘額：';

const totalAmount = document.createElement('div');
totalAmount.className = 'total-amount';
totalAmount.textContent = '---';

const updateTotalColor = (value) => {
  const numValue = parseFloat(value) || 0;
  totalAmount.classList.remove('positive', 'negative');
  totalTitle.classList.remove('positive', 'negative');
  if (numValue > 0) {
    totalAmount.classList.add('positive');
    totalTitle.classList.add('positive');
  } else if (numValue < 0) {
    totalAmount.classList.add('negative');
    totalTitle.classList.add('negative');
  }
};


// 移除圓餅圖區域

const submitContainer = document.createElement('div');
submitContainer.style.width = '100%';
submitContainer.style.display = 'flex';
submitContainer.style.justifyContent = 'center';
submitContainer.style.padding = '0';



totalContainer.appendChild(monthIndicator);
totalContainer.appendChild(columnsContainer);
      columnsContainer.appendChild(incomeColumn);
        incomeColumn.appendChild(incomeTitle);
        incomeColumn.appendChild(incomeAmount);
      columnsContainer.appendChild(expenseColumn);
        expenseColumn.appendChild(expenseTitle);
        expenseColumn.appendChild(expenseAmount);
      columnsContainer.appendChild(totalColumn);
        totalColumn.appendChild(totalTitle);
        totalColumn.appendChild(totalAmount);

budgetCardsContainer.appendChild(itemContainer);
budgetCardsContainer.appendChild(submitContainer);
  submitContainer.appendChild(mainSaveButton);

mainSaveButton.addEventListener('click', saveData);

document.addEventListener('DOMContentLoaded', async function() {
  // 顯示進度條載入下拉選單（第一個請求）
  showSpinner();

  // 一進到頁面就發送 Create Tab
  try {
    await callAPI({ name: "Create Tab" });
    // 發送 Create Tab 後，重新載入月份列表
    await loadMonthNames();
    setToIDB('sheetNames', sheetNames).catch(() => {});
    setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});
    // 如果月份選擇下拉選單已經打開，更新它
    const existingModal = document.querySelector('.month-select-modal');
    if (existingModal) {
      // 如果月份選擇彈出視窗已經打開，重新載入月份列表並更新選項
      const select = existingModal.querySelector('#month-select');
      if (select) {
        select.innerHTML = '';
        sheetNames.forEach((month, index) => {
          const option = document.createElement('option');
          // 修正：直接使用陣列索引計算 sheetIndex
          // sheetNames 已跳過前兩個無效項目（空白表、下拉選單）
          // 所以 sheetIndex = index + 2
          const sheetIndex = index + 2;

          option.value = sheetIndex;
          option.textContent = month;
          option.dataset.monthName = month;
          if (sheetIndex === currentSheetIndex) {
            option.selected = true;
          }
          select.appendChild(option);
        });
      }
    }
  } catch (e) {
    // 建立失敗，忽略錯誤（可能已經存在）
  }

  // ===== 新的載入流程：先從 IndexedDB 快取載入，再背景同步 =====
  let loadedFromCache = false;
  let cachedTimestamp = null;

  try {
    // 嘗試從 IndexedDB 載入快取資料
    const cachedSheetNames = await getFromIDB('sheetNames');
    const cachedHasInvalid = await getFromIDB('hasInvalidFirstTwoSheets');
    const cachedBudgetTotals = await getFromIDB('budgetTotals');

    if (cachedSheetNames && cachedSheetNames.length > 0) {
      sheetNames = cachedSheetNames;
      hasInvalidFirstTwoSheets = cachedHasInvalid || false;

      // 根據目前日期，設定對應的 sheet（找最接近的月份）
      const closestSheetIndex = findClosestMonth();
      currentSheetIndex = closestSheetIndex;

      // 載入當前月份的快取資料
      const cachedMonthData = await getFromIDB(`monthData_${currentSheetIndex}`);
      cachedTimestamp = await getCacheTimestamp(`monthData_${currentSheetIndex}`);

      if (cachedMonthData) {
        allMonthsData[currentSheetIndex] = cachedMonthData;

        // 載入預算快取 - 但驗證它有正確的月份資料
        if (cachedBudgetTotals && cachedBudgetTotals[currentSheetIndex]) {
          const monthName = sheetNames[currentSheetIndex - 2] || '';
          // 清除過期快取 - 強制從 API 重新載入
          // budgetTotals[currentSheetIndex] = cachedBudgetTotals[currentSheetIndex];
        }

        loadedFromCache = true;
      }
    }
  } catch (e) {
    // IndexedDB 不可用，繼續使用 API 載入
  }

  try {
    // 確保 post-content 元素存在
    const postContentElements = document.getElementsByClassName('post-content');
    if (!postContentElements || postContentElements.length === 0) {
      // 嘗試等待一下再重試
      setTimeout(() => {
        const retryElements = document.getElementsByClassName('post-content');
        if (retryElements && retryElements.length > 0) {
          retryElements[0].appendChild(totalContainer);
          retryElements[0].appendChild(budgetCardsContainer);
        } else {
        }
      }, 500);
      return;
    }

    const postContent = postContentElements[0];
    // 添加總計容器和表單容器
    postContent.appendChild(totalContainer);
    postContent.appendChild(budgetCardsContainer);

    // 第一個請求：載入下拉選單選項（阻塞式，確保表單可用）
    try {
      await loadDropdownOptions();
      updateSelectOptions('expense-category-select', EXPENSE_CATEGORY_OPTIONS);
      updateSelectOptions('payment-method-select', PAYMENT_METHOD_OPTIONS);
      updateSelectOptions('credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);
      updateSelectOptions('month-payment-select', MONTH_PAYMENT_OPTIONS);
      updateSelectOptions('payment-platform-select', PAYMENT_PLATFORM_OPTIONS);
    } catch (err) {
      console.error('[支出表] 載入下拉選單失敗:', err);
    }

    // 下拉選單載入完成，隱藏進度條
    hideSpinner();

    // 非阻塞重新載入下拉選單的函數（用於監聽更新）
    const refreshDropdowns = () => {
      loadDropdownOptions().then(() => {
        updateSelectOptions('expense-category-select', EXPENSE_CATEGORY_OPTIONS);
        updateSelectOptions('payment-method-select', PAYMENT_METHOD_OPTIONS);
        updateSelectOptions('credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);
        updateSelectOptions('month-payment-select', MONTH_PAYMENT_OPTIONS);
        updateSelectOptions('payment-platform-select', PAYMENT_PLATFORM_OPTIONS);
      }).catch(err => {
      });
    };

    // 監聽設定頁的更新通知（當設定頁更新下拉選單後，自動重新載入）
    // 使用 capture 階段確保能捕獲到事件
    window.addEventListener('storage', (e) => {
      if (e.key === 'dropdownUpdated') {
        refreshDropdowns();
      }
    }, true);

    // 監聽同頁面的自定義事件（當同一頁面觸發更新時）
    window.addEventListener('dropdownUpdated', () => {
      refreshDropdowns();
    });

    // 也監聽同頁面的 storage 事件（因為 storage 事件只在其他標籤頁觸發）
    let lastUpdateTime = localStorage.getItem('dropdownUpdated');
    const checkInterval = setInterval(() => {
      const current = localStorage.getItem('dropdownUpdated');
      if (current && current !== lastUpdateTime) {
        lastUpdateTime = current;
        refreshDropdowns();
      }
    }, 500); // 每500毫秒檢查一次，更頻繁地檢查

    // 頁面卸載時清理定時器
    window.addEventListener('beforeunload', () => {
      clearInterval(checkInterval);
    });

    // 將歷史紀錄按鈕添加到頁面標題右側（留空隙）
    const pageTitle = document.querySelector('h1.post-title, h1, .post-title');

    if (pageTitle) {
      pageTitle.style.cssText = 'display: flex; align-items: center; justify-content: space-between;';
      // 在標題和按鈕之間留空隙
      const spacer = document.createElement('div');
      spacer.style.cssText = 'flex: 1;';
      pageTitle.appendChild(spacer);
      pageTitle.appendChild(historyButton);
    } else {
      // 如果找不到標題，創建一個標題容器
      const titleContainer = document.createElement('div');
      titleContainer.style.cssText = 'display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;';
      const title = document.createElement('h1');
      title.textContent = '支出';
      titleContainer.appendChild(title);
      const spacer = document.createElement('div');
      spacer.style.cssText = 'flex: 1;';
      titleContainer.appendChild(spacer);
      titleContainer.appendChild(historyButton);
      postContent.insertBefore(titleContainer, postContent.firstChild);
    }

    // ===== 快取優先載入邏輯 =====
    if (loadedFromCache) {
      try {
        // 從快取立即顯示資料
        processDataFromResponse(allMonthsData[currentSheetIndex].data, true);
        updateTotalDisplay();

        // 背景同步：從 API 載入最新資料
        syncFromAPI().catch(e => {
          console.error('[支出表] 背景同步失敗:', e);
        });
      } catch (cacheError) {
        // 快取載入失敗，清除快取標記，讓下面的 else 分支處理
        loadedFromCache = false;
      }
    }

    // 監聽手動同步請求（從導航列的同步圖示觸發）
    window.addEventListener('syncRequested', () => {
      syncFromAPI().catch(e => {
        console.error('[支出表] 手動同步失敗:', e);
      });
    });

    if (!loadedFromCache) {
      // 沒有快取，從 API 載入
      try {
        await loadMonthNames();

        // 儲存 sheetNames 到 IndexedDB
        setToIDB('sheetNames', sheetNames).catch(() => {});
        setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});

        // 一開始就顯示總數
        const totalMonths = sheetNames.length;
        const totalProgress = totalMonths + 1;
        updateProgress(0, totalProgress, '載入月份列表');
        updateProgress(1, totalProgress, '載入月份列表');

        // 檢查當前月份是否已有表格
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const currentMonthStr = `${year}${month}`;
        const hasCurrentMonth = sheetNames.includes(currentMonthStr);

        if (!hasCurrentMonth) {
          try {
            const createResult = await callAPI({ name: "Create Tab" });
            alert(createResult.message || `已建立新分頁：${currentMonthStr}`);
            await loadMonthNames();
            setToIDB('sheetNames', sheetNames).catch(() => {});
            const newTotalMonths = sheetNames.length;
            const newTotalProgress = newTotalMonths + 1;
            updateProgress(1, newTotalProgress, '載入月份列表');
          } catch (createError) {
            alert('無法自動建立本月表格，請稍後再試或手動建立。');
          }
        }

        if (sheetNames.length > 0) {
          const closestSheetIndex = findClosestMonth();
          currentSheetIndex = closestSheetIndex;

          updateProgress(2, totalProgress, '載入當前月份');
          const currentMonthData = await loadMonthData(currentSheetIndex);
          allMonthsData[currentSheetIndex] = currentMonthData;

          // 儲存到 IndexedDB
          setToIDB(`monthData_${currentSheetIndex}`, currentMonthData).catch(() => {});

          processDataFromResponse(currentMonthData.data, true);
          updateTotalDisplay();

          // 載入預算
          loadBudgetForMonth(currentSheetIndex)
            .then(() => {
              updateTotalDisplay();
              // 儲存預算到 IndexedDB
              setToIDB('budgetTotals', budgetTotals).catch(() => {});
            })
            .catch(err => {
              console.error('[支出表] 載入預算資料失敗:', err);
            });

          // 背景預載其他月份
          preloadAllMonthsData()
            .then(() => {
              updateTotalDisplay();
              // 儲存所有月份資料到 IndexedDB
              Object.keys(allMonthsData).forEach(idx => {
                setToIDB(`monthData_${idx}`, allMonthsData[idx]).catch(() => {});
              });
            })
            .catch(e => {
              console.error('[支出表] 預載所有月份支出資料失敗:', e);
            });

          // 背景預載其他月份的預算（並行載入）
          const budgetPromises = sheetNames
            .map((name, idx) => idx + 2)
            .filter(sheetIndex => sheetIndex !== currentSheetIndex)
            .map(sheetIndex => loadBudgetForMonth(sheetIndex).catch(() => {}));

          Promise.all(budgetPromises).then(() => {
            setToIDB('budgetTotals', budgetTotals).catch(() => {});
          });
        }
      } catch (e) {
        console.error('[支出表] 初始化失敗:', e);
      }
    }

    // 等待 DOM 更新後再初始化表單，但避免清掉使用者已經輸入的內容
    setTimeout(() => {
      const itemInput = document.getElementById('item-input');
      const actualCostInput = document.getElementById('actual-cost-input');
      const recordCostInput = document.getElementById('record-cost-input');
      const noteInput = document.getElementById('note-input');

      const hasUserInput = !!(
        (itemInput && itemInput.value && itemInput.value.trim() !== '') ||
        (actualCostInput && actualCostInput.value && actualCostInput.value.trim() !== '') ||
        (recordCostInput && recordCostInput.value && recordCostInput.value !== '') ||
        (noteInput && noteInput.value && noteInput.value.trim() !== '')
      );

      if (!hasUserInput) {
        // 只有在表單目前是空的時候才執行初始化，避免清掉使用者正在輸入的內容
      clearForm();
      } else {
      }
    }, 100);

    // 載入總計
    try {
      await loadTotal();
    } catch (e) {
    }
  } catch (error) {
  }
});
</script>
