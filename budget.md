---
layout: page
title: 預算表
permalink: /budget_table/
---

<link rel="stylesheet" href="{{ '/assets/common.css' | relative_url }}?v={{ site.time | date: '%s' }}">
<link rel="stylesheet" href="{{ '/assets/budget-table.css' | relative_url }}?v={{ site.time | date: '%s' }}">

<div id="user-info"></div>

<script>

// Budget
const baseBudget = "https://script.google.com/macros/s/AKfycbxkOKU5YxZWP1XTCFCF7a62Ar71fUz4Qw7tjF3MvMGkLTt6QzzhGLnDsD7wVI_cgpAR/exec";

// ===== Global State Variables =====
let currentSheetIndex = 2; // Currently selected spreadsheet tab index (2 represents the third tab)
let sheetNames = []; // All month tab names (excluding first two: "blank sheet", "dropdown")
let allMonthsData = {}; // Pre-loaded data for all months (key: sheetIndex, value: { data, total })
let allRecords = []; // Records for currently selected month (income + expense)
let filteredRecords = []; // Records filtered by current type
let currentRecordIndex = 0;
let isNewMode = false; // Whether in add new mode
let currentRecordNumber = null; // Currently displayed record number
let isSwitchingMonth = false; // Prevent rapid consecutive month switching
let monthSelectChangeHandler = null; // Store month select event handler for removal
let currentAbortController = null; // Used to cancel ongoing requests
let timeUpdateInterval = null; // Used to update time in add new mode
let hasUnsavedChanges = false; // Track if there are unsaved changes
let originalValues = null; // Store original values for comparison

// ===== Using Shared Cache Module (SyncStatus) =====
// Using SyncStatus module's cache functionality (defined in assets/sync-status.js)
const getFromCache = async (key) => {
  try {
    return await SyncStatus.getFromCache(key);
  } catch (e) {
    return null;
  }
};

const setToCache = async (key, value) => {
  try {
    await SyncStatus.setToCache(key, value);
  } catch (e) {
    // Cache may be unavailable, ignore error
  }
};

const removeFromCache = async (key) => {
  try {
    await SyncStatus.setToCache(key, null);
  } catch (e) {
    // Cache may be unavailable, ignore error
  }
};

// ===== Dropdown Options (loaded from "dropdown" sheet=1) =====
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

// Load latest options from "dropdown" sheet=1
// API URL（支出表）
const baseExpense = "https://script.google.com/macros/s/AKfycbxpBh0QVSVTjylhh9cj7JG9d6aJi7L7y6pQPW88EbAsNtcd5ckucLagH8XpSAGa8IZt/exec";

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
      // Find the first key whose value is an array
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
    // Find corresponding column (budget table uses "expense-item")
    const colCategory = findHeaderColumn(headerRow, ['支出－項目', '支出-項目', '消費類別', '類別']);
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
      // If currently showing expense category, re-render
      const categorySelect = document.getElementById('category-select');
      if (categorySelect && categorySelect.value === '預計支出') {
        // Save current selected value (if any)
        const currentCategorySelect = document.getElementById('expense-category-select');
        const currentValue = currentCategorySelect ? currentCategorySelect.value : '';
        updateDivVisibility('預計支出');
        // If there was a previous selected value, try to restore it
        if (currentValue) {
          setTimeout(() => {
            const newCategorySelect = document.getElementById('expense-category-select');
            if (newCategorySelect && newCategorySelect.querySelector(`option[value="${currentValue}"]`)) {
              newCategorySelect.value = currentValue;
              // Sync update custom dropdown display text
              const selectContainer = newCategorySelect.parentElement;
              if (selectContainer) {
                const selectDisplay = selectContainer.querySelector('.select-display');
                if (selectDisplay) {
                  const selectText = selectDisplay.querySelector('.select-text');
                  if (selectText) {
                    const selectedOpt = newCategorySelect.options[newCategorySelect.selectedIndex];
                    selectText.textContent = selectedOpt ? selectedOpt.textContent : '';
                  }
                }
              }
            }
          }, 100);
        }
      }
    } else {
    }

  } catch (err) {
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

// ===== Unified API Call Function =====
async function callAPI(postData) {
  // Cancel previous request (if exists)
  if (currentAbortController) {
    currentAbortController.abort();
  }

  // Create new AbortController
  currentAbortController = new AbortController();
  const signal = currentAbortController.signal;

  try {
    const response = await fetch(baseBudget, {
      method: "POST",
      redirect: "follow",
      mode: "cors",
      keepalive: true,
      signal: signal, // Add abort signal
      body: JSON.stringify(postData)
    });

    const responseText = await response.text();
    if (!responseText || responseText.trim() === '') {
      currentAbortController = null;
      return { success: true, data: null, total: null };
    }

    let result;
    try {
      result = JSON.parse(responseText);
    } catch (e) {
      currentAbortController = null;
      throw new Error('後端響應格式錯誤: ' + responseText.substring(0, 100));
    }

    if (!response.ok || !result.success) {
      currentAbortController = null;
      throw new Error(result.message || result.error || '操作失敗');
    }

    // After successful request, clear cache to ensure latest data on next load
    removeFromCache(`budget_monthData_${currentSheetIndex}`);
    delete allMonthsData[currentSheetIndex];

    // After successful request, clear AbortController
    currentAbortController = null;
    return result;
  } catch (error) {
    // If request was cancelled, don't throw error
    if (error.name === 'AbortError') {
      currentAbortController = null;
      throw new Error('Request cancelled');
    }
    currentAbortController = null;
    throw error;
  }
}

// ===== Find the Closest Month (current month or latest month) =====
function findClosestMonth() {
  if (sheetNames.length === 0) return 2;

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const currentMonthStr = `${currentYear}${String(currentMonth).padStart(2, '0')}`;

  const currentIndex = sheetNames.findIndex(name => name === currentMonthStr);
  if (currentIndex !== -1) {
    return currentIndex + 2;
  }

  return sheetNames.length + 1; // sheetIndex of the last month
}

// ===== Enter Add New Mode if No Records =====
function enterNewModeIfEmpty() {
  if (filteredRecords.length > 0) return;

  const currentType = categorySelect ? categorySelect.value : '預計支出';
  let nextNumber = 1;
  const allRecordsOfType = allRecords.filter(r => r.type === currentType);
  if (allRecordsOfType.length > 0) {
    const maxNum = Math.max(
      ...allRecordsOfType
        .map(r => parseInt(r.row[0], 10))
        .filter(n => Number.isFinite(n) && n > 0)
    );
    if (Number.isFinite(maxNum) && maxNum > 0) {
      nextNumber = maxNum + 1;
    }
  }
  isNewMode = true;
  if (typeof recordNumber !== 'undefined') {
    recordNumber.textContent = ''; // Don't show number in add new mode
    recordNumber.style.display = 'none'; // Hide number
  }
  if (typeof recordDate !== 'undefined') {
    recordDate.textContent = getNowFormattedDateTime();
  }

  // Start time update timer (updates every second)
  if (timeUpdateInterval) {
    clearInterval(timeUpdateInterval);
  }
  timeUpdateInterval = setInterval(() => {
    if (isNewMode && typeof recordDate !== 'undefined') {
      recordDate.textContent = getNowFormattedDateTime();
    } else {
      // Stop timer when not in add new mode
      if (timeUpdateInterval) {
        clearInterval(timeUpdateInterval);
        timeUpdateInterval = null;
      }
    }
  }, 1000);
  const itemInput = document.getElementById('item-input');
  const costInput = document.getElementById('cost-input');
  const noteInput = document.getElementById('note-input');
  if (itemInput) itemInput.value = '';
  if (costInput) costInput.value = '';
  if (noteInput) noteInput.value = '';
  updateDeleteButton();
  updateArrowButtons();
  // Store empty values as original values (add new mode)
  storeOriginalValues();
}

// Filter records by currently selected type
function filterRecordsByType(type) {
  filteredRecords = allRecords.filter(r => r.type === type);

  // Ensure currentRecordIndex is within valid range
  if (currentRecordIndex >= filteredRecords.length && filteredRecords.length > 0) {
    currentRecordIndex = filteredRecords.length - 1;
  } else if (filteredRecords.length === 0) {
    currentRecordIndex = 0;
  }

  // Add new mode: switching income/expense mode: recalculate next number for this type
  if (isNewMode) {
    // Calculate max number + 1 in new type
    let nextNumber = 1;
    if (filteredRecords.length > 0) {
      const maxNum = Math.max(
        ...filteredRecords
          .map(r => parseInt(r.row[0], 10))
          .filter(n => Number.isFinite(n) && n > 0)
      );
      if (Number.isFinite(maxNum) && maxNum > 0) {
        nextNumber = maxNum + 1;
      }
    }

    // Don't show number in add new mode
    if (typeof recordNumber !== 'undefined') {
      recordNumber.textContent = ''; // Don't show number in add new mode
      recordNumber.style.display = 'none'; // Hide number
    }

    // Clear form, prepare for new entry
    const itemInput = document.getElementById('item-input');
    const costInput = document.getElementById('cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.value = '';
    if (costInput) costInput.value = '';
    if (noteInput) noteInput.value = '';

    // If expense, reset category selection
    if (type === '預計支出') {
      const categorySelectElement = document.getElementById('expense-category-select');
      if (categorySelectElement && categorySelectElement.options.length > 0) {
        categorySelectElement.value = categorySelectElement.options[0].value;
        const selectContainer = categorySelectElement.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('div');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('div');
            if (selectText) {
              selectText.textContent = categorySelectElement.options[0].textContent;
            }
          }
        }
      }
    }

    updateArrowButtons();
    updateDeleteButton(); // Update delete button display
    return;
  }

  // Try to find record with same number
  if (currentRecordNumber !== null && filteredRecords.length > 0) {
    const sameNumberIndex = filteredRecords.findIndex(r => {
      const num = parseInt(r.row[0], 10);
      return Number.isFinite(num) && num > 0 && num === currentRecordNumber;
    });

    if (sameNumberIndex >= 0) {
      currentRecordIndex = sameNumberIndex;
      showRecord(sameNumberIndex);
      updateArrowButtons();
      return;
    }
  }

  // If same number not found, show first record
  currentRecordIndex = 0;
  if (filteredRecords.length > 0) {
    showRecord(0);
  } else {
    // When no records, enter new mode so adding works correctly
    enterNewModeIfEmpty();
  }
}

// Get current time and format as YYYY/MM/DD HH:MM
function getNowFormattedDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  return `${year}/${month}/${day} ${hours}:${minutes}`;
}

// Convert various time string formats to YYYY/MM/DD HH:MM
function formatRecordDateTime(raw) {
  if (!raw) return '';

  // Try to parse as Date (supports ISO format like 2025-11-30T12:34:56Z)
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) {
    // On parse failure, keep original string (e.g., already 2025/11/30 12:34)
    return raw;
  }

  const year = dt.getFullYear();
  const month = String(dt.getMonth() + 1).padStart(2, '0');
  const day = String(dt.getDate()).padStart(2, '0');
  const hours = String(dt.getHours()).padStart(2, '0');
  const minutes = String(dt.getMinutes()).padStart(2, '0');
  return `${year}/${month}/${day} ${hours}:${minutes}`;
}

// Display record at index to card input fields
function showRecord(index) {
  if (!filteredRecords.length) return;

  // Ensure index is within valid range
  if (index < 0 || index >= filteredRecords.length) {
    index = Math.max(0, Math.min(index, filteredRecords.length - 1));
  }

  currentRecordIndex = index; // Update current index
  const { type, row } = filteredRecords[index];
  isNewMode = false; // Exit add new mode when showing record

  // Stop time update timer
  if (timeUpdateInterval) {
    clearInterval(timeUpdateInterval);
    timeUpdateInterval = null;
  }

  updateDeleteButton(); // Update delete button display
  markAsSaved(); // Reset to unchanged state when showing record

  // Update top-left number: only show the record's own number (row[0]), not income/expense
  // Ensure number element exists and displays correctly
  let recordNumberEl = document.getElementById('record-number');
  
  // If element not found, try using global variable
  if (!recordNumberEl && typeof recordNumber !== 'undefined' && recordNumber) {
    recordNumberEl = recordNumber;
  }
  
  if (recordNumberEl) {
    const num = parseInt(row[0], 10);
    const recordNum = Number.isFinite(num) && num > 0 ? num : (index + 1);
    const numberText = `#${String(recordNum).padStart(3, '0')}`;
    
    // Force set number content and style (use !important to override CSS display: none)
    recordNumberEl.textContent = numberText;
    recordNumberEl.setAttribute('style', 'display: block !important; visibility: visible !important; opacity: 1 !important;');
    recordNumberEl.removeAttribute('hidden'); // Remove hidden attribute
    
    // Record current number
    currentRecordNumber = recordNum;
    
    // Debug log
  } else {
  }

  // Update top-right "data time": use each record's time field (row[1], usually timestamp/last modified time)
  if (typeof recordDate !== 'undefined') {
    recordDate.textContent = formatRecordDateTime(row[1] || '');
  }

  // Set "expense/income" category (only change display, don't trigger change event to avoid recursion)
  if (typeof categorySelect !== 'undefined' && typeof categorySelectText !== 'undefined') {
    const value = type === '收入' || type === '預計支出' ? type : '預計支出';
    categorySelect.value = value;
    categorySelectText.textContent = value;
    // Directly update field display
    if (typeof updateDivVisibility === 'function') {
      updateDivVisibility();
    }
  }

  // Wait for div2/div3/div4 to be built based on category before filling data
  // Delay 200ms to ensure updateDivVisibility's clone operation (100ms) is complete
  // And wait for DOM to stabilize before setting values
  setTimeout(() => {
    // Function to set input field values
    const setInputValues = () => {
      const itemInput = document.getElementById('item-input');
      const costInput = document.getElementById('cost-input');
      const noteInput = document.getElementById('note-input');

      if (type === '預計支出') {
        // Expense: [number, time, category, item, cost, note]
        const categorySelectElement = document.getElementById('expense-category-select');
        if (categorySelectElement) {
          categorySelectElement.value = row[2] || '';
          // Sync update custom dropdown display text (if exists)
          const selectContainer = categorySelectElement.parentElement;
          if (selectContainer) {
            const selectDisplay = selectContainer.querySelector('div');
            if (selectDisplay) {
              const selectText = selectDisplay.querySelector('div');
              if (selectText) {
                selectText.textContent = row[2] || '';
              }
            }
          }
        }
        if (itemInput) {
          itemInput.value = row[3] || '';
          // Ensure value is set (avoid placeholder showing)
          if (itemInput.value === '' && row[3]) {
            itemInput.value = row[3];
          }
        }
        // Ensure amount displays correctly (convert to number then back to string to avoid format issues)
        if (costInput) {
          const costValue = row[4];
          // Handle same as income: support displaying number 0
          if (costValue !== undefined && costValue !== null && costValue !== '') {
            const numCost = parseFloat(costValue);
            costInput.value = Number.isFinite(numCost) ? numCost.toString() : '';
          } else if (costValue === 0 || costValue === '0') {
            // Show even if 0
            costInput.value = '0';
          } else {
            costInput.value = '';
          }
        }
        if (noteInput) noteInput.value = row[5] || '';
      } else {
        // Income: [number, time, item, cost, note]
        if (itemInput) {
          itemInput.value = row[2] || '';
          // Ensure value is set (avoid placeholder showing)
          if (itemInput.value === '' && row[2]) {
            itemInput.value = row[2];
          }
        }
        // Ensure amount displays correctly (convert to number then back to string to avoid format issues)
        if (costInput) {
          const costValue = row[3];
          // Income amount in row[3], show even if 0 or empty
          // Ensure all cases display correctly
          if (costValue !== undefined && costValue !== null && costValue !== '') {
            const numCost = parseFloat(costValue);
            const finalValue = Number.isFinite(numCost) ? numCost.toString() : String(costValue);
            costInput.value = finalValue;
          } else if (costValue === 0 || costValue === '0') {
            // Show even if 0
            costInput.value = '0';
          } else {
            // If empty string or null/undefined, show empty string
            costInput.value = '';
          }
        }
        if (noteInput) noteInput.value = row[4] || '';
      }

      // Update arrow button states
      updateArrowButtons();
      
      // Store original values (after input fields are set)
      storeOriginalValues();
    };
    
    // Try setting values first
    setInputValues();
    
    // Delay another 50ms to set values, ensure correct display after clone operation
    // This handles cases where updateDivVisibility's clone operation (100ms) may complete late
    setTimeout(setInputValues, 50);
  }, 200);
}

// Find record with specified number and type from data
const findRecordInData = (data, recordNumber, recordType) => {
  if (!data || typeof data !== 'object') return null;
  
  const monthIndex = currentSheetIndex - 2;
  const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : '';
  
  const isIncome = recordType === '收入';
  const isExpense = recordType === '預計支出';
  
  // Find corresponding named range
  for (const key of Object.keys(data)) {
    const rows = data[key] || [];
    const keyIsIncome = key.includes('收入');
    const keyIsExpense = key.includes('支出');
    
    // Check if type matches
    if ((isIncome && !keyIsIncome) || (isExpense && !keyIsExpense)) {
      continue;
    }
    
    // Check if month matches
    if (currentMonthName && !key.includes(currentMonthName)) {
      continue;
    }
    
    // Find record with matching number in rows
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;
      
      const num = parseInt(row[0], 10);
      if (Number.isFinite(num) && num > 0 && num === recordNumber) {
        // Found corresponding record
        const expectedLength = isIncome ? 5 : 6;
        const processedRow = [...row]; // Copy array
        if (processedRow.length < expectedLength) {
          while (processedRow.length < expectedLength) {
            processedRow.push('');
          }
        }
        return { type: recordType, row: processedRow };
      }
    }
  }
  
  return null;
};

// Process data returned from Apps Script (for updating allRecords)
const processDataFromResponse = (data, shouldFilter = true) => {
  // Debug: log loaded data
  const monthIndex = currentSheetIndex - 2;
  const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : 'unknown';
  console.log(`[Budget] Loading month: ${currentMonthName} (sheetIndex: ${currentSheetIndex})`);
  console.log('[Budget] Raw data:', data);
  
  // First clear current records
  allRecords = [];

  if (!data) {
    return;
  }

  // Use Set to track processed records, avoid duplicates
  const processedRecords = new Set();
  let processedCount = 0;
  let totalRowsCount = 0;

  // Budget table only processes object format data (named range format)
  // Named ranges: CurrentMonthIncome202506, CurrentMonthExpenseBudget202506
  // Budget format: Income [number, time, item, cost, note] (5 cols), Expense [number, time, category, item, cost, note] (6 cols)
  if (data && typeof data === 'object' && !Array.isArray(data)) {
    // Process object format data (named range format)
    Object.keys(data).forEach(key => {
      const rows = data[key] || [];
      totalRowsCount += rows.length;

      const isIncome = key.includes('收入');
      const isExpense = key.includes('支出');
      const type = isIncome ? '收入' : (isExpense ? '預計支出' : '');

      if (!type) {
        return;
      }

      // Only process current month's data (named range name should contain current month)
      // If month name exists, must match; if not, process all matching type data (for initialization)
      let isCurrentMonth = false;
      if (!currentMonthName || currentMonthName === '') {
        // If no month name, process all matching type data (might be initialization stage)
        isCurrentMonth = true;
      } else {
        // Check if key contains current month name
        isCurrentMonth = key.includes(currentMonthName);
        // If direct match fails, try checking if it contains month number (e.g., 202506)
        if (!isCurrentMonth && currentMonthName.length >= 6) {
          // Try extracting month number part (e.g., from "202506" extract "202506")
          const monthNum = currentMonthName.substring(0, 6); // Get first 6 chars (YYYYMM)
          isCurrentMonth = key.includes(monthNum);
        }
        // If still fails, try matching only year and month (e.g., "2025" + "06")
        if (!isCurrentMonth && currentMonthName.length >= 6) {
          const year = currentMonthName.substring(0, 4);
          const month = currentMonthName.substring(4, 6);
          isCurrentMonth = key.includes(year) && key.includes(month);
        }
      }
      if (!isCurrentMonth) {
        return;
      }

      rows.forEach((row, rowIndex) => {
        if (!row || !Array.isArray(row) || row.length === 0) {
          return;
        }

        // First check if first column is a valid number (priority check to avoid wrongly skipping data rows)
        const firstCell = row[0];
        // Try multiple ways to parse number
        let num = parseInt(firstCell, 10);
        if (isNaN(num)) {
          // If parseInt fails, try parseFloat
          num = parseFloat(firstCell);
        }
        const isValidNumber = Number.isFinite(num) && num > 0;
        
        // Debug: log processing status of each row
        if (!isValidNumber) {
        }

        // If valid number, treat as data row (even if column count is insufficient or other checks might misjudge)
        if (isValidNumber) {
          // Budget format check: income should be 5 cols, expense should be 6 cols
          // But if number is valid, accept even if column count is insufficient (might be partial data)
          const expectedLength = type === '收入' ? 5 : 6;
          if (row.length < expectedLength) {
            // If columns insufficient, pad to expected length
            while (row.length < expectedLength) {
              row.push('');
            }
          }

          // Use number+time as unique identifier to avoid duplicates
          const recordKey = `${type}_${num}_${row[1] || ''}`;
          if (processedRecords.has(recordKey)) {
            return;
          }
          processedRecords.add(recordKey);

          allRecords.push({ type, row });
          processedCount++;
          return; // Skip subsequent checks, continue to next row
        }

        // If not valid number, check column count (might be incomplete data)
        // Budget format check: income should be 5 cols, expense should be 6 cols
        // But allow slightly relaxed column count (some columns might be empty)
        if (type === '收入' && row.length < 5) {
          return;
        }
        if (type === '預計支出' && row.length < 6) {
          return;
        }

        // If not valid number, also check if it's header or total row
        const firstCellStr = String(firstCell || '').trim();
        const firstCellLower = firstCellStr.toLowerCase();
        // Check if header row (contains common header keywords)
        if (firstCellLower === '交易日期' || firstCellLower === '編號' || firstCellLower === '日期' ||
            firstCellLower === '時間' || firstCellLower === '總計' || firstCellStr === '' ||
            firstCellLower.includes('項目') || firstCellLower.includes('金額') ||
            firstCellLower.includes('備註') || firstCellLower.includes('類別')) {
          return;
        }

        // If neither valid number nor header/total row, might be malformed data row
        // To avoid missing data, if first column is not empty string, try parsing as number
        if (firstCellStr !== '') {
          const altNum = parseFloat(firstCellStr);
          if (Number.isFinite(altNum) && altNum > 0) {
            // Pad columns to expected length
            const expectedLength = type === '收入' ? 5 : 6;
            if (row.length < expectedLength) {
              while (row.length < expectedLength) {
                row.push('');
              }
            }
            const recordKey = `${type}_${altNum}_${row[1] || ''}`;
            if (!processedRecords.has(recordKey)) {
              processedRecords.add(recordKey);
              allRecords.push({ type, row });
              processedCount++;
            }
          }
        }
      });
    });
    
    // Debug: log processed data
    console.log(`[Budget] Processing complete, total ${allRecords.length} records`);
    if (allRecords.length > 0) {
      console.log('[Budget] Record list:', allRecords.map(r => ({
        type: r.type,
        number: r.row[0],
        time: r.row[1],
        category: r.type === '預計支出' ? r.row[2] : 'income',
        item: r.type === '預計支出' ? r.row[3] : r.row[2],
        amount: r.type === '預計支出' ? r.row[4] : r.row[3]
      })));
    }
  } else if (Array.isArray(data)) {
    // If data is array format, try converting to object format
    // This case should be rare since budget table uses named range format
    // But for robustness, we still handle it
    // Can try converting array to object, but need to know data structure
    // Skip for now since budget table should always return object format
  } else {
  }

  // Filter records by currently selected type (default show expense)
  if (shouldFilter) {
    const currentType = categorySelect ? categorySelect.value : '預計支出';
    filterRecordsByType(currentType);
  }
};

// If totalData is null or undefined, calculate live total from current records
const updateTotalDisplay = (totalData = null) => {
  let income = 0;
  let expense = 0;
  let total = 0;
  // Important: Regardless of backend returned total, recalculate from allRecords to ensure correct calculation when editing amount (total - old value + new value)
  // Because when editing amount, allRecords is already updated to new value, so calculating from allRecords gives correct total
  if (false && totalData && Array.isArray(totalData) && totalData.length >= 3) {
    // Temporarily not using backend returned total, calculate from allRecords instead
    // Use backend returned total
    // Validate if each value is valid number
    const incomeRaw = totalData[0];
    const expenseRaw = totalData[1];
    const totalRaw = totalData[2];
    
    income = parseFloat(incomeRaw);
    expense = parseFloat(expenseRaw);
    total = parseFloat(totalRaw);
    
    // If parse fails, use 0
    if (isNaN(income)) {
      income = 0;
    }
    if (isNaN(expense)) {
      expense = 0;
    }
    if (isNaN(total)) {
      total = income - expense; // Use calculated value
    }
  } else {
    // Calculate live total from current records
    // Important: Initially sum all, only exclude old value and add new value when user modifies amount
    const costInput = document.getElementById('cost-input');
    const currentType = categorySelect ? categorySelect.value : '預計支出';
    
    // Check if user is editing amount (edit mode and costInput.value differs from current record amount)
    let isEditingAmount = false;
    let currentRecord = null;
    let oldAmount = 0;
    
    if (!isNewMode && filteredRecords.length > 0 && currentRecordIndex >= 0 && currentRecordIndex < filteredRecords.length) {
      currentRecord = filteredRecords[currentRecordIndex];
      const currentRecordNumber = parseInt(currentRecord.row[0], 10);
      
      // Get old amount of current record
      if (currentRecord.type === '收入') {
        oldAmount = parseFloat(currentRecord.row[3] || 0) || 0;
      } else {
        oldAmount = parseFloat(currentRecord.row[4] || 0) || 0;
      }
      
      // Check if user is editing amount
      if (costInput && costInput.value !== undefined && costInput.value !== null && costInput.value !== '') {
        const newAmount = parseFloat(costInput.value) || 0;
        // If new amount differs from old amount, user is editing
        if (Math.abs(newAmount - oldAmount) > 0.01) {
          isEditingAmount = true;
        }
      }
    }
    
    // If user is editing amount, need to exclude old record (old value), then add new value (costInput.value)
    // Otherwise, sum all (including current record)
    let recordsToCalculate = allRecords;
    if (isEditingAmount && currentRecord) {
      const currentRecordNumber = parseInt(currentRecord.row[0], 10);
      
      // Exclude currently editing record (old value) from allRecords
      recordsToCalculate = allRecords.filter(r => {
        const num = parseInt(r.row[0], 10);
        return !(Number.isFinite(num) && num > 0 && num === currentRecordNumber && r.type === currentRecord.type);
      });
    } else {
      // Initially sum all (including current record), don't add costInput.value
      recordsToCalculate = allRecords;
    }
    
    const incomeRecords = recordsToCalculate.filter(r => r.type === '收入');
    const expenseRecords = recordsToCalculate.filter(r => r.type === '預計支出');
    // Calculate income total
    income = incomeRecords.reduce((sum, r) => {
      if (!r || !r.row || !Array.isArray(r.row)) return sum;
      const cost = parseFloat(r.row[3] || 0) || 0; // Income: row[3] is amount
      return sum + cost;
    }, 0);

    // Calculate expense total
    expense = expenseRecords.reduce((sum, r) => {
      if (!r || !r.row || !Array.isArray(r.row)) return sum;
      const cost = parseFloat(r.row[4] || 0) || 0; // Expense: row[4] is amount
      return sum + cost;
    }, 0);

    // Add live input amount (new value)
    // Add new mode: add costInput.value (because new value not yet in allRecords)
    // Edit mode and editing amount: add costInput.value (because old value already excluded from allRecords)
    // Edit mode but not editing amount: don't add costInput.value (because value already in allRecords)
    if (isNewMode || isEditingAmount) {
      if (costInput && costInput.value) {
        const liveCost = parseFloat(costInput.value) || 0;
        if (currentType === '收入') {
          income += liveCost;
        } else {
          expense += liveCost;
        }
      }
    } else {
    }

    total = income - expense;
  }

  incomeAmount.textContent = income.toLocaleString('zh-TW');
  expenseAmount.textContent = expense.toLocaleString('zh-TW');
  totalAmount.textContent = total.toLocaleString('zh-TW');
  updateTotalColor(total);
};

// Load single month's data and total
const loadMonthData = async (sheetIndex, useGlobalAbortController = true) => {
  // Validate if sheetIndex is valid
  if (!Number.isFinite(sheetIndex) || sheetIndex < 2) {
    throw new Error(`Invalid sheet index: ${sheetIndex}`);
  }

  let signal;
  let abortController;

  if (useGlobalAbortController) {
    // Cancel previous request (if exists)
    if (currentAbortController) {
      currentAbortController.abort();
    }

    // Create new AbortController
    currentAbortController = new AbortController();
    signal = currentAbortController.signal;
  } else {
    // Create independent AbortController for preload tasks
    abortController = new AbortController();
    signal = abortController.signal;
  }

  // Fetch "current month income/expense" data from spreadsheet - add timestamp to avoid cache
  const monthIndex = sheetIndex - 2;
  const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : '';

  const dataParams = { name: "Show Tab Data", sheet: sheetIndex, _t: Date.now() };
  const dataUrl = `${baseBudget}?${new URLSearchParams(dataParams)}`;

  let data;
  try {
    const res = await fetch(dataUrl, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store", // Force no cache
      signal: signal // Add abort signal
    });

    if (!res.ok) {
      throw new Error(`載入資料失敗: HTTP ${res.status} ${res.statusText}`);
    }

    data = await res.json();

    // Display returned data in console
    // If data is array format, need to convert to named range format
    if (Array.isArray(data)) {

      // Convert array to named range format
      // Budget format: Income [number, time, item, cost, note] (5 cols)
      //          Expense [number, time, category, item, cost, note] (6 cols)
      const convertedData = {};
      let incomeRows = [];
      let expenseRows = [];

      data.forEach((row, rowIndex) => {
        if (!row || !Array.isArray(row) || row.length === 0) return;

        // Skip header row and total row
        const firstCell = String(row[0] || '').trim();
        if (firstCell === '交易日期' || firstCell === '編號' || firstCell === '總計' || firstCell === '') {
          return;
        }

        // Determine type based on column count
        if (row.length === 5) {
          // Income format: [number, time, item, cost, note]
          incomeRows.push(row);
        } else if (row.length === 6) {
          // Expense format: [number, time, category, item, cost, note]
          expenseRows.push(row);
        } else {
          // Other formats (like 10-col expense records) skip, budget table only needs budget data
        }
      });

      // Convert to named range format
      if (incomeRows.length > 0) {
        convertedData[`當月收入${currentMonthName}`] = incomeRows;
      }
      if (expenseRows.length > 0) {
        convertedData[`當月支出預算${currentMonthName}`] = expenseRows;
      }

      data = convertedData;
    } else {
      // Object format (named range format), check if named range exists
      const expectedIncomeKey = `當月收入${currentMonthName}`;
      const expectedExpenseKey = `當月支出預算${currentMonthName}`;
      const hasIncome = data.hasOwnProperty(expectedIncomeKey);
      const hasExpense = data.hasOwnProperty(expectedExpenseKey);
    }
  } catch (error) {
    if (error.name === 'AbortError') {
      throw new Error(`載入月份 ${currentMonthName} (sheetIndex: ${sheetIndex}) 被取消`);
    }
    throw error;
  }

  // Load total - add timestamp to avoid cache
  const TotalParams = { name: "Show Total", sheet: sheetIndex, _t: Date.now() };
  const Totalurl = `${baseBudget}?${new URLSearchParams(TotalParams)}`;
  const Totalres = await fetch(Totalurl, {
    method: "GET",
    redirect: "follow",
    mode: "cors",
    cache: "no-store", // Force no cache
    signal: signal // Add abort signal
  });

  if (!Totalres.ok) {
    throw new Error(`載入總計失敗: HTTP ${Totalres.status} ${Totalres.statusText}`);
  }

  let totalData;
  try {
    totalData = await Totalres.json();
  } catch (e) {
    throw new Error('解析總計資料失敗: ' + e.message);
  }

  // Validate total data format
  if (!Array.isArray(totalData) || totalData.length < 3) {
    // If format error, use [0, 0, 0] as default value
    totalData = [0, 0, 0];
  }

  // Validate if total data values are valid numbers
  const income = parseFloat(totalData[0]);
  const expense = parseFloat(totalData[1]);
  const total = parseFloat(totalData[2]);
  
  if (isNaN(income) || isNaN(expense) || isNaN(total)) {
    // If contains invalid number, use 0 as default value
    totalData = [
      Number.isFinite(income) ? income : 0,
      Number.isFinite(expense) ? expense : 0,
      Number.isFinite(total) ? total : 0
    ];
  }

  // Display returned total data in console
  // After successful request, clear AbortController (only when using global AbortController)
  if (useGlobalAbortController) {
    currentAbortController = null;
  }

  // Calculate total from data (because Google Apps Script total may have issues)
  let calculatedIncome = 0;
  let calculatedExpense = 0;
  let incomeCount = 0;
  let expenseCount = 0;

  // Use Set to track processed records, avoid duplicate calculations
  const processedIncomeRecords = new Set();
  const processedExpenseRecords = new Set();

  if (data && typeof data === 'object') {
    // Debug: check 202506 month data filtering
    // Based on new Apps Script, data format is:
    // key is named range name (e.g., "CurrentMonthIncome202506", "CurrentMonthExpenseBudget202506")
    // value is 2D array (empty rows filtered)
    Object.keys(data).forEach(key => {
      const rows = data[key] || [];

      // Determine type based on named range name
      // Named range format: CurrentMonthIncome202506 or CurrentMonthExpenseBudget202506
      const isIncome = key.includes('收入');
      const isExpense = key.includes('支出');

      if (!isIncome && !isExpense) {
        return; // Skip data that is not income or expense
      }

      // Only process current month's data (named range name should contain current month)
      const isCurrentMonth = currentMonthName && key.includes(currentMonthName);
      if (!isCurrentMonth) {
        return; // Skip data that is not current month
      }

      rows.forEach((row, rowIndex) => {
        // Ensure row is array
        if (!row || !Array.isArray(row) || row.length === 0) return;

        // Check if empty row (all columns empty)
        const isEmptyRow = row.every(cell => cell === '' || cell === null || cell === undefined);
        if (isEmptyRow) {
          return;
        }

        // Skip total row and header row (stricter check)
        const firstCell = String(row[0] || '').trim();
        const firstCellLower = firstCell.toLowerCase();
        // Check if header row or total row
        if (firstCellLower === '交易日期' || firstCellLower === '總計' || firstCell === '' ||
            firstCell === null || firstCell === undefined ||
            firstCellLower.includes('項目') || firstCellLower.includes('金額') ||
            firstCellLower.includes('備註') || firstCellLower.includes('類別') ||
            firstCellLower === '編號' || firstCellLower === '日期' || firstCellLower === '時間') {
          return;
        }

        // Check if valid record (first column should be number)
        const num = parseInt(firstCell, 10);
        if (!Number.isFinite(num) || num <= 0) {
          return;
        }

        // Check if this record already processed (use number+time as unique identifier)
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

        // Income: [number, time, item, cost, note] - cost at index 3 (column D)
        // Expense: [number, time, category, item, cost, note] - cost at index 4 (column K)
        const costIndex = isIncome ? 3 : (isExpense ? 4 : -1);
        if (costIndex >= 0 && row[costIndex] !== undefined && row[costIndex] !== null && row[costIndex] !== '') {
          const cost = parseFloat(row[costIndex]);
          if (Number.isFinite(cost) && cost !== 0) { // Allow negative, but don't sum 0
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

  // Save raw data directly, no format conversion needed
  // Apps Script returns standard format (e.g., "CurrentMonthIncome202506", "CurrentMonthExpenseBudget202506")
  // Keep as-is when saving, match by month name when reading

  // Validate if API returned total format is correct, prefer calculated value (especially when calculated value is not 0)
  let finalTotal = calculatedTotalData; // Default use calculated total
  
  if (totalData && Array.isArray(totalData) && totalData.length >= 3) {
    const apiIncome = parseFloat(totalData[0]) || 0;
    const apiExpense = parseFloat(totalData[1]) || 0;
    const apiTotal = parseFloat(totalData[2]) || 0;
    
    // If calculated has expense records (expenseCount > 0), prefer calculated value
    if (expenseCount > 0 || calculatedExpense > 0) {
      // Has expense records, use calculated value
      finalTotal = calculatedTotalData;
      if (Math.abs(apiExpense - calculatedExpense) > 0.01) {
      }
    } else if (calculatedIncome > 0 || calculatedExpense > 0) {
      // Calculated has income or expense, use calculated value
      finalTotal = calculatedTotalData;
    } else if (apiExpense === 0 && apiIncome === 0) {
      // Both are 0, use API value (might be more accurate)
      finalTotal = totalData;
    } else {
      // No calculated value, but API has value, use API value
      finalTotal = totalData;
    }
  } else {
    // API returned total format incorrect, use calculated total
  }

  // Display final returned data and total in console
  return { data, total: finalTotal };
};

// Preload other months' data (in order: next month first, then backwards)
const preloadAllMonthsData = async (baseProgress = 2, totalProgress = 0) => {
  if (sheetNames.length === 0) {
    return;
  }

  // Find current month's index in sheetNames
  const currentMonthIdx = sheetNames.findIndex((name, idx) => {
    const sheetIndex = idx + 2;
    return sheetIndex === currentSheetIndex;
  });
  if (currentMonthIdx === -1) {
    return;
  }

  // Build load order: current month first (if not loaded), then next month and after, finally previous months
  const loadOrder = [];

  // First check if current month loaded, if not add to load order
  const currentMonthSheetIndex = currentMonthIdx + 2;
  if (!allMonthsData[currentMonthSheetIndex]) {
    loadOrder.push({ idx: currentMonthIdx, sheetIndex: currentMonthSheetIndex, name: sheetNames[currentMonthIdx] });
  } else {
  }

  // First add next month and subsequent months (starting from currentMonthIdx + 1)
  for (let i = currentMonthIdx + 1; i < sheetNames.length; i++) {
    const sheetIndex = i + 2;
    if (!allMonthsData[sheetIndex]) {
      loadOrder.push({ idx: i, sheetIndex, name: sheetNames[i] });
    } else {
    }
  }

  // Then add previous months backwards (from currentMonthIdx - 1)
  for (let i = currentMonthIdx - 1; i >= 0; i--) {
    const sheetIndex = i + 2;
    if (!allMonthsData[sheetIndex]) {
      loadOrder.push({ idx: i, sheetIndex, name: sheetNames[i] });
    } else {
    }
  }
  if (loadOrder.length === 0) {
    return;
  }

  let loadedCount = 0;

  // Load in order (one by one, not concurrent)
  for (const item of loadOrder) {
    // First check if pre-fetched data already exists
    if (allMonthsData[item.sheetIndex]) {
      loadedCount++;
      if (totalProgress > 0) {
        updateProgress(baseProgress + loadedCount, totalProgress, `載入月份 ${item.name}（從快取）`);
      }
      continue; // Skip, use existing data
    }

    try {
      // Use independent AbortController to avoid conflicts with user operations
      const monthData = await loadMonthData(item.sheetIndex, false);
      allMonthsData[item.sheetIndex] = monthData;

      // Save to cache
      setToCache(`budget_monthData_${item.sheetIndex}`, monthData);
      // Update progress bar
      loadedCount++;
      if (totalProgress > 0) {
        updateProgress(baseProgress + loadedCount, totalProgress, `載入月份 ${item.name}`);
      }

      // If encounter incomplete, exit (show progress bar)
      // Here we continue loading, but progress bar keeps updating
    } catch (error) {
      // Update progress even if failed
      loadedCount++;
      if (totalProgress > 0) {
        updateProgress(baseProgress + loadedCount, totalProgress, `載入月份 ${item.name}`);
      }

      // If error occurs, continue loading next month (don't interrupt)
    }
  }
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
      const storedData = await getFromCache(`budget_monthData_${currentSheetIndex}`);
      if (storedData) {
        monthData = storedData;
        // Explicitly mark as loaded from cache
        monthData._fromCache = true;
        // Also load to memory
        allMonthsData[currentSheetIndex] = monthData;
      } else {
      }
    } catch (e) {
      // Cache may be unavailable or data corrupted, ignore error
    }
  }

  if (!monthData) {
    return false; // Indicates need to reload
  }
  // When loading from storage, show loading and block header
  showSpinner(true);

  // Process data (will auto filter and display records)
  if (monthData.data) {
    processDataFromResponse(monthData.data, true);
    // Update total immediately after processing data (ensure total auto-calculated)
    if (monthData.total && Array.isArray(monthData.total) && monthData.total.length >= 3) {
      updateTotalDisplay(monthData.total);
    } else {
      // If no total or format incorrect, use live calculation
      updateTotalDisplay();
    }
  } else {
    allRecords = [];
    filteredRecords = [];
    // Even if no data, update total (display as 0)
    updateTotalDisplay();
  }
  // Ensure first record displayed (if any, and not in add new mode)
  if (filteredRecords.length > 0) {
    showRecord(0);
    updateArrowButtons();
  } else {
    // If no records, enter add new mode
    enterNewModeIfEmpty();
    updateDeleteButton();
    updateArrowButtons();
  }

  // Hide spinner after loading complete
  setTimeout(() => {
    hideSpinner();
  }, 100);

  // Reset button state after loading complete
  markAsSaved();

  return true; // Indicates successfully loaded from memory
};

// Load current month's data (prefer memory, send request if not available)
const loadContent = async (forceReload = false) => {
  // If not forcing reload, try reading from memory first
  if (!forceReload) {
    const loadedFromMemory = await loadContentFromMemory();
    if (loadedFromMemory) {
      return; // Successfully loaded from memory, return directly
    }
  } else {
  }

  // If no data in memory, or need to force reload, send request
  try {
    const monthData = await loadMonthData(currentSheetIndex);
    // Update data in memory
    allMonthsData[currentSheetIndex] = monthData;

    // Save to cache
    setToCache(`budget_monthData_${currentSheetIndex}`, monthData);
    // Process data (will auto filter and display records)
    if (monthData.data) {
      processDataFromResponse(monthData.data, true);
      // Update total immediately after processing data (ensure total auto-calculated)
      if (monthData.total && Array.isArray(monthData.total) && monthData.total.length >= 3) {
        updateTotalDisplay(monthData.total);
      } else {
        // If no total or format incorrect, use live calculation
        updateTotalDisplay();
      }
    } else {
      allRecords = [];
      filteredRecords = [];
      // Even if no data, update total (display as 0)
      updateTotalDisplay();
    }

    // Ensure first record displayed (if any, and not in add new mode)
    if (filteredRecords.length > 0) {
      showRecord(0);
      updateArrowButtons();
    } else {
      // If no records, enter add new mode
      enterNewModeIfEmpty();
      updateDeleteButton();
      updateArrowButtons();
    }
    
    // Reset button state after loading complete
    markAsSaved();
  } catch (error) {
    throw error;
  }
};


const loadTotal = async () => {
  // Validate if currentSheetIndex is valid
  if (!Number.isFinite(currentSheetIndex) || currentSheetIndex < 2) {
    currentSheetIndex = 2; // Default to third tab
  }

  // Prefer reading total from memory
  const monthData = allMonthsData[currentSheetIndex];
  if (monthData && monthData.total) {
    updateTotalDisplay(monthData.total);
    return;
  }

  // If not in memory, send request
  try {
    const TotalParams = { name: "Show Total", sheet: currentSheetIndex };
  const Totalurl = `${baseBudget}?${new URLSearchParams(TotalParams)}`;
    const Totalres = await fetch(Totalurl, {
      method: "GET",
      redirect: "follow",
      mode: "cors"
    });

    if (!Totalres.ok) {
      throw new Error(`Failed to load total: HTTP ${Totalres.status} ${Totalres.statusText}`);
    }

  const Totaldata = await Totalres.json();

    // Update total in memory (if data exists)
    if (allMonthsData[currentSheetIndex]) {
      allMonthsData[currentSheetIndex].total = Totaldata;
    } else {
      // If data doesn't exist, create a new entry
      allMonthsData[currentSheetIndex] = { total: Totaldata };
    }

    // Save to cache
    setToCache(`budget_monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]);

    updateTotalDisplay(Totaldata);
  } catch (error) {
    // Don't throw error, just log, avoid affecting other functions
  }
};

// Timer for progress bar auto animation
let progressAnimationTimer = null;
let progressAnimationStartTime = null;
let progressAnimationTarget = 99; // Auto animation target percentage (99%)

// Update progress bar (actual progress, overrides auto animation)
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

// Start progress bar auto animation (0% to 99% in 10 seconds)
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

    // If reached 99% or exceeded, stop animation
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

  // Create fullscreen overlay (coverHeader decides whether to cover header)
  const overlay = document.createElement('div');
  overlay.id = 'loading-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: ${coverHeader ? '0' : '60px'};
    left: 0;
    right: 0;
    bottom: 0;
    background: linear-gradient(135deg, rgba(232, 248, 240, 0.97) 0%, rgba(240, 255, 245, 0.97) 50%, rgba(230, 255, 240, 0.97) 100%);
    z-index: ${coverHeader ? '2000' : '1500'}; /* If covering header, z-index should be higher than header */
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

  // Start auto progress bar animation (0% to 99% in 10 seconds)
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

    // Slightly delay before removing overlay to let user see 100%
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
  
  // Add input change listener to check for actual changes
  input.addEventListener('input', () => {
    checkForChanges();
  });

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
  
  // Add input change listener to check for actual changes
  textarea.addEventListener('input', () => {
    checkForChanges();
  });

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

  const selectText = document.createElement('div');
  selectText.className = 'select-text';
  selectText.textContent = options[0].text;

  const selectArrow = document.createElement('div');
  selectArrow.className = 'select-arrow';
  selectArrow.textContent = '▼';

  selectDisplay.appendChild(selectText);
  selectDisplay.appendChild(selectArrow);

  const hiddenSelect = document.createElement('select');
  hiddenSelect.id = selectId;
  hiddenSelect.name = selectId; // Add name attribute to support autofill
  hiddenSelect.style.display = 'none';
  hiddenSelect.value = options[0].value;

  const dropdown = document.createElement('div');
  dropdown.className = 'select-dropdown';

  options.forEach(opt => {
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
      checkForChanges(); // Check for actual changes
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
    
    // Close all other dropdowns before opening this one
    if (!isOpen) {
      // Close category dropdown
      const categoryDropdown = document.querySelector('.category-dropdown');
      if (categoryDropdown) {
        categoryDropdown.style.display = 'none';
        const categoryArrow = document.querySelector('.category-select-arrow');
        if (categoryArrow) {
          categoryArrow.style.transform = 'rotate(0deg)';
        }
      }
      
      // Close all other select dropdowns
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
    }
    
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

const updateDivVisibility = (forceType = null) => {
  // If type parameter provided, use it; otherwise try to get latest element value from DOM
  let categoryValue = forceType;
  if (categoryValue === null) {
    // First try to get from global variable
    if (typeof categorySelect !== 'undefined' && categorySelect.value) {
      categoryValue = categorySelect.value;
    } else {
      // If global variable not available, get latest element from DOM
      const categorySelectElement = document.getElementById('category-select');
      if (categorySelectElement) {
        categoryValue = categorySelectElement.value;
      } else {
        categoryValue = '預計支出'; // Default value
      }
    }
  }

  div2.innerHTML = '';
  div3.innerHTML = '';
  div4.innerHTML = '';

  if (categoryValue === '預計支出') {
    const categoryRow = createSelectRow('類別：', 'expense-category-select', EXPENSE_CATEGORY_OPTIONS);
    const costRow = createInputRow('金額：', 'cost-input', 'number');
    const noteRow = createTextareaRow('備註：', 'note-input', 3);
    noteRow.style.marginBottom = '0px';

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

    div2.appendChild(costRow);
    div3.appendChild(noteRow);

    itemContainer.style.display = 'flex';
    div2.style.display = 'flex';
    div3.style.display = 'flex';
    div4.style.display = 'none';
  }

  // Add listener for real-time total update for cost input field
  setTimeout(() => {
    const costInput = document.getElementById('cost-input');
    if (costInput) {
      // Save current value (if any)
      const currentValue = costInput.value;
      
      // Remove old listener (if any) to avoid duplicate additions
      const newCostInput = costInput.cloneNode(true);
      costInput.parentNode.replaceChild(newCostInput, costInput);
      
      // Restore previous value (cloneNode should have preserved it, but set again for safety)
      if (currentValue) {
        newCostInput.value = currentValue;
      }

      // Add new listener
      newCostInput.addEventListener('input', () => {
        updateTotalDisplay(); // No params, use live calculation
        checkForChanges(); // Check for actual changes
      });
    }
  }, 100);
};

const saveData = async () => {
  const categoryValue = categorySelect.value;
  const itemInput = document.getElementById('item-input');
  const costInput = document.getElementById('cost-input');
  const noteInput = document.getElementById('note-input');
  const TotalParams = { name: "Show Total", sheet: currentSheetIndex };

  if (!itemInput || !costInput) {
    alert('請等待表單載入完成');
    return;
  }

  const item = itemInput.value.trim();
  const costValue = costInput.value.trim();
  const cost = parseFloat(costValue);
  const note = noteInput ? noteInput.value.trim() : '';

  if (!item) {
    alert('請輸入項目');
    return;
  }

  if (!costValue || isNaN(cost) || cost <= 0) {
    alert('請輸入有效金額');
    return;
  }

  let category = '';
  let range = 0;

  if (categoryValue === '預計支出') {
    const div2Display = window.getComputedStyle(div2).display;
    if (div2Display === 'none' || div2Display === '') {
      alert('請等待表單載入完成');
      return;
    }
    const categorySelectElement = document.getElementById('expense-category-select');
    if (!categorySelectElement) {
      alert('請等待表單載入完成');
      return;
    }
    if (!categorySelectElement.value) {
      alert('請選擇支出類別');
      return;
    }
    category = categorySelectElement.value;
    range = 0;
  } else {
    range = 1;
  }


  // Lock entire page, wait for backend response
  showSpinner();
  saveButton.textContent = '儲存中...';
  saveButton.disabled = true;
  saveButton.style.opacity = '0.6';
  saveButton.style.cursor = 'not-allowed';

  // Disable all inputs and buttons
  if (itemInput) itemInput.disabled = true;
  if (costInput) costInput.disabled = true;
  if (noteInput) noteInput.disabled = true;
  if (categorySelect) categorySelect.disabled = true;
  const expenseCategorySelect = document.getElementById('expense-category-select');
  if (expenseCategorySelect) expenseCategorySelect.disabled = true;
  leftArrow.disabled = true;
  rightArrow.disabled = true;
  deleteButton.disabled = true;
  if (monthSelect) monthSelect.disabled = true;

  let alreadyReset = false; // Ensure button state only restored at appropriate time
  try {
    const postData = {
      name: "Upsert Data",
      sheet: currentSheetIndex,
      range: range,
      item: item,
      cost: cost,
      note: note,

    };

    // For expense (range === 0), category must be sent
    if (range === 0) {
      postData.category = category;
    }

    // If not in add new mode, must send updateRow parameter to update existing record
    // updateRow is row number = record number + 1 (first row is header, data starts from second row)
    if (!isNewMode) {
      if (filteredRecords.length > 0 && currentRecordIndex < filteredRecords.length) {
        const currentRecord = filteredRecords[currentRecordIndex];
        const recordNum = parseInt(currentRecord.row[0], 10);
        console.log('[Budget] Update record - recordNum:', recordNum, 'row[0]:', currentRecord.row[0], 'updateRow:', recordNum + 2);
        if (Number.isFinite(recordNum) && recordNum > 0) {
          postData.updateRow = recordNum + 2; // Row 1: type header, Row 2: field header, Row 3+: data
        } else {
          alert('無法更新記錄：找不到記錄編號');
          return;
        }
      } else {
        alert('無法更新記錄：找不到目前記錄');
        return;
      }
    }


    const response = await fetch(baseBudget, {
      method: "POST",
      redirect: "follow",
      mode: "cors",
      keepalive: true,
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify(postData)
    });

    const responseText = await response.text();

    let result;
    try {
      result = JSON.parse(responseText);
    } catch (e) {
      throw new Error('Backend response format error: ' + responseText);
    }

    if (response.ok && result.success) {
      // Record whether in add new mode
      const wasInNewMode = isNewMode;
      const savedType = categoryValue;
      // Record current editing record number (for edit mode)
      const savedRecordNumber = !wasInNewMode && filteredRecords.length > 0 && currentRecordIndex < filteredRecords.length
        ? parseInt(filteredRecords[currentRecordIndex].row[0], 10)
        : null;

      alert('資料已成功儲存！');

      // Before reloading, set currentRecordNumber first (for edit mode)
      if (!wasInNewMode && savedRecordNumber !== null) {
        currentRecordNumber = savedRecordNumber;
      }

      // If Apps Script returned data and total, use directly, otherwise reload
      if (result.data && result.total) {
        // If in edit mode, remove old record from allRecords first (avoid double total)
        if (!wasInNewMode && savedRecordNumber !== null) {
          // Find and remove old record
          const oldRecordIndex = allRecords.findIndex(r => {
            const num = parseInt(r.row[0], 10);
            return Number.isFinite(num) && num > 0 && num === savedRecordNumber;
          });
          if (oldRecordIndex >= 0) {
            allRecords.splice(oldRecordIndex, 1);
          } else {
          }
        }

        // Update current month's data in memory
        allMonthsData[currentSheetIndex] = {
          data: result.data,
          total: result.total
        };

        // Save to cache
        setToCache(`budget_monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]);

        // Record currently displayed record number (for repositioning after update)
        const currentDisplayedRecordNumber = !wasInNewMode && filteredRecords.length > 0 && currentRecordIndex < filteredRecords.length
          ? parseInt(filteredRecords[currentRecordIndex].row[0], 10)
          : null;

        // Use returned data to update records and total (don't auto filter, filter manually later)
        // Important: whether adding or editing, reprocess all data to ensure no missing records
        // Because backend has correctly handled total (delete old value then recalculate), frontend should also reprocess all data
        processDataFromResponse(result.data, false);
        
        // Important: must exit add new mode before calculating total
        // Because allRecords already contains new record, if isNewMode is still true,
        // updateTotalDisplay will add costInput.value again, causing total to double
        isNewMode = false;
        
        // Ensure total correctly updated
        // Recalculate total from allRecords (allRecords is already latest data)
        updateTotalDisplay(null); // Pass null to force calculation from allRecords

        // Re-filter records based on saved type
        const currentType = categorySelect ? categorySelect.value : savedType;
        filterRecordsByType(currentType);

        // Wait for filtering to complete
        await new Promise(r => setTimeout(r, 100));

        // If was in edit mode, try to find and position to the edited record
        if (!wasInNewMode && currentDisplayedRecordNumber !== null) {
          const recordIndex = filteredRecords.findIndex(r => {
            const num = parseInt(r.row[0], 10);
            return Number.isFinite(num) && num > 0 && num === currentDisplayedRecordNumber;
          });

          if (recordIndex >= 0) {
            currentRecordIndex = recordIndex;
            showRecord(recordIndex);
            updateArrowButtons();
          } else {
            // If not found, show first record
            if (filteredRecords.length > 0) {
              currentRecordIndex = 0;
              showRecord(0);
              updateArrowButtons();
            }
          }
        } else if (wasInNewMode) {
          // Add new mode: find last record (just added) and display
          if (filteredRecords.length > 0) {
            // Find last record (should be newly added)
            const lastIndex = filteredRecords.length - 1;
            currentRecordIndex = lastIndex;
            showRecord(lastIndex);
            updateArrowButtons();
        } else {
            // If no records, enter add new mode
            currentRecordIndex = 0;
            enterNewModeIfEmpty();
            updateArrowButtons();
          }
        } else {
          // Edit mode but record not found, show first record
          currentRecordIndex = 0;
          if (filteredRecords.length > 0) {
            showRecord(0);
            updateArrowButtons();
          }
        }
      } else {
      // Reload records (avoid cache, add short delay)
      try {
        await new Promise(r => setTimeout(r, 200));
        await loadContent();
        } catch (e) {
        }
      }

        // If was in add new mode, show the newly added record
        if (wasInNewMode) {
          // Ensure type is correct
          if (categorySelect && categorySelect.value !== savedType) {
            categorySelect.value = savedType;
            categorySelectText.textContent = savedType;
            filterRecordsByType(savedType);
          }

          // After filtering completes, find and show the newly added record
          setTimeout(() => {
            if (filteredRecords.length > 0) {
              // Find last record (should be newly added)
              const lastIndex = filteredRecords.length - 1;
              currentRecordIndex = lastIndex;
              showRecord(lastIndex);
            updateArrowButtons();
            } else {
              // If no records, enter add new mode
              enterNewModeIfEmpty();
              updateArrowButtons();
            }
          }, 150);
        } else {
          // If not add new mode (edit mode), re-display the edited record
          if (savedRecordNumber !== null) {
            // After filtering completes, find and show the edited record
            // Use longer delay to ensure loadContent and filterRecordsByType complete
            setTimeout(() => {
              const recordIndex = filteredRecords.findIndex(r => {
                const num = parseInt(r.row[0], 10);
                return Number.isFinite(num) && num > 0 && num === savedRecordNumber;
              });

              if (recordIndex >= 0) {
                currentRecordIndex = recordIndex;
                showRecord(recordIndex);
                updateArrowButtons();
              } else {
                // If not found, show first record
                if (filteredRecords.length > 0) {
                  currentRecordIndex = 0;
                  showRecord(0);
                  updateArrowButtons();
                }
              }
            }, 300);
          } else {
            // If record number not found, show first record
            setTimeout(() => {
              if (filteredRecords.length > 0) {
                currentRecordIndex = 0;
                showRecord(0);
                updateArrowButtons();
              }
            }, 100);
          }
        }

      // Restore button state after total update completes
      saveButton.textContent = '儲存';
      markAsSaved(); // Mark as saved, disable button
      storeOriginalValues(); // Store current values as new original values
      alreadyReset = true;
    } else {
      const errorMessage = result.message || result.error || '未知錯誤';
      alert('儲存失敗: ' + errorMessage);
      // On failure, restore button state (keep enabled due to unsaved changes)
      saveButton.textContent = '儲存';
      if (hasUnsavedChanges) {
        saveButton.disabled = false;
        saveButton.style.opacity = '1';
        saveButton.style.cursor = 'pointer';
      } else {
        markAsSaved(); // If no changes, disable button
      }
      alreadyReset = true;
    }
  } catch (error) {
    alert('儲存失敗: ' + error.message);
    // On exception, restore button state
    saveButton.textContent = '儲存';
    if (hasUnsavedChanges) {
      saveButton.disabled = false;
      saveButton.style.opacity = '1';
      saveButton.style.cursor = 'pointer';
    } else {
      markAsSaved(); // If no changes, disable button
    }
    alreadyReset = true;
  } finally {
    // Restore all buttons and inputs
    hideSpinner();
    if (!alreadyReset) {
      saveButton.textContent = '儲存';
      if (hasUnsavedChanges) {
        saveButton.disabled = false;
        saveButton.style.opacity = '1';
        saveButton.style.cursor = 'pointer';
      } else {
        markAsSaved(); // If no changes, disable button
      }
    }

    // Restore all inputs and buttons
    if (itemInput) itemInput.disabled = false;
    if (costInput) costInput.disabled = false;
    if (noteInput) noteInput.disabled = false;
    if (categorySelect) categorySelect.disabled = false;
    const expenseCategorySelect = document.getElementById('expense-category-select');
    if (expenseCategorySelect) expenseCategorySelect.disabled = false;
    leftArrow.disabled = false;
    rightArrow.disabled = false;
    deleteButton.disabled = false;
    if (monthSelect) monthSelect.disabled = false;
  }
};


const totalContainer = document.createElement('div');
totalContainer.className = 'total-container';

const budgetCardsContainer = document.createElement('div');
budgetCardsContainer.className = 'budget-cards-container';

const headerInfo = document.createElement('div');
headerInfo.className = 'header-info';

const recordNumber = document.createElement('div');
recordNumber.id = 'record-number';
recordNumber.textContent = '#001';
// Use !important to override CSS display: none
recordNumber.setAttribute('style', 'display: block !important; visibility: visible !important;');

const recordDate = document.createElement('div');
recordDate.id = 'record-date';
recordDate.textContent = ''; // Display each record's time (e.g., last modified time in spreadsheet)

// recordNumber is now directly appended to budgetCardsContainer (during DOM construction)
// This way its position: absolute; left: 15px; can be correctly positioned
headerInfo.appendChild(recordDate);

// Place month dropdown to the right of page title (e.g., right of "Budget Table")
document.addEventListener('DOMContentLoaded', () => {
  const titleEl = document.querySelector('.post-title');
  if (titleEl && monthSelectWrapper) {
    titleEl.appendChild(monthSelectWrapper);
  }
});

// Delete button
const deleteButton = document.createElement('button');
deleteButton.className = 'delete-button';
deleteButton.textContent = '刪除';

// Function to delete current record (can be called by button and keyboard)
const deleteCurrentRecord = async () => {
  // Cannot delete in add new mode
  if (isNewMode) {
    alert('無法刪除：目前為新增模式');
    return;
  }

  // Confirm deletion
  if (!confirm('確定要刪除這筆記錄嗎？')) {
    return;
  }

  if (!filteredRecords.length || currentRecordIndex >= filteredRecords.length) {
    alert('無法刪除：找不到目前記錄');
    return;
  }

  const currentRecord = filteredRecords[currentRecordIndex];
  const recordNum = parseInt(currentRecord.row[0], 10);
  const recordType = currentRecord.type;

  if (!Number.isFinite(recordNum) || recordNum <= 0) {
    alert('無法刪除：找不到記錄編號');
    return;
  }

  // Determine rangeType (0=expense, 1=income)
  const rangeType = recordType === '預計支出' ? 0 : 1;

  // Lock entire page, wait for backend response
  showSpinner();
  deleteButton.disabled = true;

  // Disable all inputs and buttons
  const itemInput = document.getElementById('item-input');
  const costInput = document.getElementById('cost-input');
  const noteInput = document.getElementById('note-input');
  if (itemInput) itemInput.disabled = true;
  if (costInput) costInput.disabled = true;
  if (noteInput) noteInput.disabled = true;
  if (categorySelect) categorySelect.disabled = true;
  const expenseCategorySelect = document.getElementById('expense-category-select');
  if (expenseCategorySelect) expenseCategorySelect.disabled = true;
  leftArrow.disabled = true;
  rightArrow.disabled = true;
  saveButton.disabled = true;
  if (monthSelect) monthSelect.disabled = true;

  try {
    const postData = {
      name: "Delete Data",
      sheet: currentSheetIndex,
      range: rangeType,
      number: recordNum.toString() // Ensure string, as Google Apps Script uses string comparison
    };

    const response = await fetch(baseBudget, {
      method: "POST",
      redirect: "follow",
      mode: "cors",
      keepalive: true,
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify(postData)
    });

    const responseText = await response.text();
    let result;
    try {
      result = JSON.parse(responseText);
    } catch (e) {
      throw new Error('Backend response format error: ' + responseText);
    }

    if (response.ok && result.success) {
      alert('記錄已成功刪除！');

      // If Apps Script returned data and total, use directly, otherwise reload
      if (result.data && result.total) {
        // Update current month's data in memory
        allMonthsData[currentSheetIndex] = {
          data: result.data,
          total: result.total
        };

        // Save to cache
        setToCache(`budget_monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]);

        // Use returned data to update records and total
        processDataFromResponse(result.data, false);
        
        // Re-filter records after deletion (ensure filteredRecords updated)
        const currentType = categorySelect ? categorySelect.value : '預計支出';
        filterRecordsByType(currentType);

        // Update total immediately (right after updating data)
        // Ensure total correctly updated (if backend returned correct format)
        if (result.total && Array.isArray(result.total) && result.total.length >= 3) {
          // Update total in memory
          allMonthsData[currentSheetIndex].total = result.total;
        updateTotalDisplay(result.total);
        } else {
          // If backend didn't return total or format incorrect, use live calculation first, then try reloading
          updateTotalDisplay(); // Immediately use live calculation to update total
          // Then try reloading total (non-blocking)
          loadTotal().catch(() => {
            // If load fails, keep using live calculation result
          });
        }
      } else {
        // Reload records (and update memory)
      await new Promise(r => setTimeout(r, 200));
        await loadContent(true); // Force reload
      }

      // After deletion, show first record (if any remain)
      await new Promise(r => setTimeout(r, 100));
        if (filteredRecords.length > 0) {
          currentRecordIndex = 0;
          showRecord(0);
          updateArrowButtons();
        } else {
          // If no records, enter add new mode
          isNewMode = true;
          updateDeleteButton(); // Update delete button display
          const itemInput = document.getElementById('item-input');
          const costInput = document.getElementById('cost-input');
          const noteInput = document.getElementById('note-input');
          if (itemInput) itemInput.value = '';
          if (costInput) costInput.value = '';
          if (noteInput) noteInput.value = '';
          if (typeof recordNumber !== 'undefined') {
          recordNumber.textContent = ''; // Don't show number in add new mode
          recordNumber.style.display = 'none'; // Hide number
          }
          // In add new mode with no data, data time uses current time
          if (typeof recordDate !== 'undefined') {
          recordDate.value = getNowFormattedDateTime();
          }
          updateArrowButtons();
        }
    } else {
      const errorMessage = result.message || result.error || '未知錯誤';
      alert('刪除失敗: ' + errorMessage);
    }
  } catch (error) {
    alert('刪除失敗: ' + error.message);
  } finally {
    // Restore all buttons and inputs
    hideSpinner();
    deleteButton.disabled = false;

    // Restore all inputs and buttons
    const itemInput = document.getElementById('item-input');
    const costInput = document.getElementById('cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.disabled = false;
    if (costInput) costInput.disabled = false;
    if (noteInput) noteInput.disabled = false;
    if (categorySelect) categorySelect.disabled = false;
    const expenseCategorySelect = document.getElementById('expense-category-select');
    if (expenseCategorySelect) expenseCategorySelect.disabled = false;
    leftArrow.disabled = false;
    rightArrow.disabled = false;
    saveButton.disabled = false;
    if (monthSelect) monthSelect.disabled = false;
  }
};

// Delete button click event
deleteButton.addEventListener('click', deleteCurrentRecord);

// Update delete button display state
function updateDeleteButton() {
  // Hide delete button in add new mode
  if (isNewMode) {
    deleteButton.style.display = 'none';
  } else {
    deleteButton.style.display = 'flex';
  }
}

const leftArrow = document.createElement('button');
leftArrow.className = 'arrow-button left';
leftArrow.innerHTML = '‹';

const rightArrow = document.createElement('button');
rightArrow.className = 'arrow-button right';
rightArrow.innerHTML = '›';

// Left/right keys switch all records
// Update arrow button state
function updateArrowButtons() {
  // If no records and is add new mode, hide all arrows
  if (!filteredRecords.length && isNewMode) {
    leftArrow.style.display = 'none';
    rightArrow.style.display = 'none';
    return;
  }

  // If no records and not in add new mode, also hide all arrows
  if (!filteredRecords.length && !isNewMode) {
    leftArrow.style.display = 'none';
    rightArrow.style.display = 'none';
    return;
  }

  // If has records, show arrows normally
  // If in add new mode, only show left arrow, hide right plus
  if (isNewMode) {
    leftArrow.style.display = 'flex';
    leftArrow.innerHTML = '‹';
    leftArrow.classList.remove('plus');
    rightArrow.style.display = 'none';
  } else {
    // If at first record, hide left arrow
    if (currentRecordIndex === 0) {
    leftArrow.style.display = 'none';
  } else {
  leftArrow.style.display = 'flex';
      leftArrow.innerHTML = '‹';
      leftArrow.classList.remove('plus');
    }

    rightArrow.style.display = 'flex';

    // If at last record, right arrow becomes plus
    if (currentRecordIndex === filteredRecords.length - 1) {
      rightArrow.innerHTML = '+';
      rightArrow.classList.add('plus');
    } else {
      rightArrow.innerHTML = '›';
      rightArrow.classList.remove('plus');
    }
  }
}

// Switch to previous record
function goToPreviousRecord() {
  // If in add new mode, return to last record
  if (isNewMode && filteredRecords.length > 0) {
    currentRecordIndex = filteredRecords.length - 1;
    showRecord(currentRecordIndex);
    updateArrowButtons();
    return;
  }

  if (!filteredRecords.length) return;

  // Ensure currentRecordIndex is within valid range
  if (currentRecordIndex >= filteredRecords.length) {
    currentRecordIndex = filteredRecords.length - 1;
  }

    currentRecordIndex = Math.max(0, currentRecordIndex - 1);
    if (currentRecordIndex < filteredRecords.length) {
    showRecord(currentRecordIndex);
    }
    updateArrowButtons();
}

// Switch to next record or enter add new mode
function goToNextRecord() {
  if (!filteredRecords.length) return;

  // Ensure currentRecordIndex is within valid range
  if (currentRecordIndex >= filteredRecords.length) {
    currentRecordIndex = filteredRecords.length - 1;
  }

  // If at last record or add new mode, enter add new mode
  if (currentRecordIndex === filteredRecords.length - 1 || isNewMode) {
    isNewMode = true; // Enter add new mode
    updateDeleteButton(); // Update delete button display

    // Clear form, prepare for adding
    const itemInput = document.getElementById('item-input');
    const costInput = document.getElementById('cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.value = '';
    if (costInput) costInput.value = '';
    if (noteInput) noteInput.value = '';

    // Calculate next number (max number in current type + 1)
    let nextNumber = 1;
    if (filteredRecords.length > 0) {
      const maxNum = Math.max(
        ...filteredRecords
          .map(r => parseInt(r.row[0], 10))
          .filter(n => Number.isFinite(n) && n > 0)
      );
      if (Number.isFinite(maxNum) && maxNum > 0) {
        nextNumber = maxNum + 1;
      }
    }

    // Don't show number in add new mode
    if (typeof recordNumber !== 'undefined') {
      recordNumber.textContent = ''; // Don't show number in add new mode
      recordNumber.style.display = 'none'; // Hide number
    }

    // "Data time" in add new mode uses current time
    if (typeof recordDate !== 'undefined') {
      recordDate.textContent = getNowFormattedDateTime();
    }

    updateArrowButtons();
  } else {
    currentRecordIndex = Math.min(filteredRecords.length - 1, currentRecordIndex + 1);
    showRecord(currentRecordIndex);
    updateArrowButtons();
  }
}

// Arrow button events
leftArrow.addEventListener('click', goToPreviousRecord);
rightArrow.addEventListener('click', goToNextRecord);

// Switch income/expense type
function switchType(targetType) {
    // Ensure categorySelect element exists
    const categorySelectElement = document.getElementById('category-select');
    if (!categorySelectElement) {
      return; // If element doesn't exist, return directly
    }

    const currentType = categorySelectElement.value || '預計支出';

    // If target type differs from current type, switch
    if (currentType !== targetType) {
      // In non-add mode, save current record number to find same number record in target type
      // Check if in add new mode: if isNewMode is true or no records, then is add new mode
      const isActuallyNewMode = isNewMode || (filteredRecords.length === 0);

      if (!isActuallyNewMode) {
        // Prefer getting number from current record
        if (filteredRecords.length > 0 && currentRecordIndex < filteredRecords.length) {
          const currentRecord = filteredRecords[currentRecordIndex];
          const recordNum = parseInt(currentRecord.row[0], 10);
          if (Number.isFinite(recordNum) && recordNum > 0) {
            currentRecordNumber = recordNum; // Save current number
          }
        }
        // If getting from record fails, try reading from displayed number element
        if (currentRecordNumber === null && typeof recordNumber !== 'undefined') {
          const recordNumText = recordNumber.textContent;
          const match = recordNumText.match(/#(\d+)/);
          if (match && match[1]) {
            const recordNum = parseInt(match[1], 10);
            if (Number.isFinite(recordNum) && recordNum > 0) {
              currentRecordNumber = recordNum;
            }
          }
        }
      }

      // First update global variable so updateDivVisibility can read correct value
      if (typeof categorySelect !== 'undefined') {
        categorySelect.value = targetType;
      }

      // Update DOM element
      categorySelectElement.value = targetType;

      // Update display text
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

      // First update UI element display (especially expense/income field switching)
      if (typeof updateDivVisibility === 'function') {
        updateDivVisibility();
      }

      // Then filter records and update UI (filterRecordsByType auto handles same number lookup)
      // Use setTimeout to ensure updateDivVisibility completes before executing
      setTimeout(() => {
      // updateDivVisibility recreates expense-category-select element, need to re-get and update
      const newCategorySelectElement = document.getElementById('expense-category-select');
        if (newCategorySelectElement) {
          // Update newly created element's value (if target type is expense, set first option; if income, no need)
          if (targetType === '預計支出' && newCategorySelectElement.options.length > 0) {
            newCategorySelectElement.value = newCategorySelectElement.options[0].value;
            // Sync update custom dropdown display text
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

          // Update global variable's value attribute (though element changed, we can keep consistency by updating attribute)
          // Actually, since categorySelect is const, we need to ensure subsequent code uses getElementById to get latest element
          // But for compatibility, we also update global variable's value (if element still exists)
          if (typeof categorySelect !== 'undefined' && categorySelect.parentNode) {
            categorySelect.value = targetType;
          }
        }

        if (typeof filterRecordsByType === 'function') {
          filterRecordsByType(targetType);
        }

        // Update related UI elements
        if (typeof updateArrowButtons === 'function') {
          updateArrowButtons();
        }
        if (typeof updateDeleteButton === 'function') {
          updateDeleteButton();
        }
      }, 200);
  }
}

// Keyboard events (computer left/right keys switch records, up/down keys switch income/expense, Delete key deletes record)
document.addEventListener('keydown', (e) => {
  // If typing text, don't trigger switching (except Delete key, as it's for deletion)
  const activeElement = document.activeElement;
  if (activeElement && (activeElement.tagName === 'INPUT' || activeElement.tagName === 'TEXTAREA')) {
    // Delete key in input field only deletes text, doesn't trigger record deletion
    if (e.key === 'Delete' || e.key === 'Backspace') {
      return;
    }
    return;
  }

  if (e.key === 'ArrowLeft') {
    e.preventDefault();
    goToPreviousRecord();
  } else if (e.key === 'ArrowRight') {
    e.preventDefault();
    goToNextRecord();
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    // Up key = switch to income
    switchType('收入');
  } else if (e.key === 'ArrowDown') {
    e.preventDefault();
    // Down key = switch to expense
    switchType('預計支出');
  } else if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    // Delete key or Backspace key = delete current record
    deleteCurrentRecord();
  }
});

// Touch swipe events (phone swipe left/right to switch records)
let touchStartX = 0;
let touchEndX = 0;
let touchStartY = 0;
let touchEndY = 0;

budgetCardsContainer.addEventListener('touchstart', (e) => {
  touchStartX = e.changedTouches[0].screenX;
  touchStartY = e.changedTouches[0].screenY;
}, { passive: true });

budgetCardsContainer.addEventListener('touchend', (e) => {
  touchEndX = e.changedTouches[0].screenX;
  touchEndY = e.changedTouches[0].screenY;
  const deltaX = touchEndX - touchStartX;
  const deltaY = touchEndY - touchStartY;

  // Handle horizontal swipe (horizontal distance clearly greater than vertical, and horizontal > 50px)
  if (Math.abs(deltaX) > Math.abs(deltaY) * 2 && Math.abs(deltaX) > 50) {
    if (deltaX > 0) {
      // Swipe right = previous
      goToPreviousRecord();
    } else {
      // Swipe left = next
      goToNextRecord();
    }
  }
}, { passive: true });

const div1 = document.createElement('div');
div1.className = 'category-select-row';

const categoryLabel = document.createElement('label');
categoryLabel.className = 'category-select-label';
categoryLabel.textContent = '類別：';
categoryLabel.htmlFor = 'category-select'; // Associate with select

const categorySelectContainer = document.createElement('div');
categorySelectContainer.className = 'category-select-container';

const categorySelectDisplay = document.createElement('div');
categorySelectDisplay.className = 'category-select-display';

const categorySelectText = document.createElement('div');
categorySelectText.className = 'category-select-text';
categorySelectText.textContent = '預計支出';

const categorySelectArrow = document.createElement('div');
categorySelectArrow.className = 'category-select-arrow';
categorySelectArrow.textContent = '▼';

categorySelectDisplay.appendChild(categorySelectText);
categorySelectDisplay.appendChild(categorySelectArrow);

const categorySelect = document.createElement('select');
categorySelect.id = 'category-select'; // Add id to support label association
categorySelect.name = 'category-select'; // Add name attribute to support autofill
categorySelect.style.display = 'none';
categorySelect.value = '預計支出';
const optionExpense = document.createElement('option');
optionExpense.value = '預計支出';
optionExpense.textContent = '預計支出';
const optionIncome = document.createElement('option');
optionIncome.value = '收入';
optionIncome.textContent = '收入';
categorySelect.appendChild(optionExpense);
categorySelect.appendChild(optionIncome);

const categoryDropdown = document.createElement('div');
categoryDropdown.className = 'category-dropdown';

const categoryOption1 = document.createElement('div');
categoryOption1.className = 'category-option';
categoryOption1.textContent = '預計支出';
categoryOption1.dataset.value = '預計支出';
categoryOption1.addEventListener('click', function() {
  categorySelectText.textContent = '預計支出';
  categorySelect.value = '預計支出';
  categoryDropdown.style.display = 'none';
  categorySelectArrow.style.transform = 'rotate(0deg)';
  categorySelect.dispatchEvent(new Event('change'));
});

const categoryOption2 = document.createElement('div');
categoryOption2.className = 'category-option';
categoryOption2.textContent = '收入';
categoryOption2.dataset.value = '收入';
categoryOption2.addEventListener('click', function() {
  categorySelectText.textContent = '收入';
  categorySelect.value = '收入';
  categoryDropdown.style.display = 'none';
  categorySelectArrow.style.transform = 'rotate(0deg)';
  categorySelect.dispatchEvent(new Event('change'));
});

categoryDropdown.appendChild(categoryOption1);
categoryDropdown.appendChild(categoryOption2);

categorySelectDisplay.addEventListener('click', function(e) {
  e.stopPropagation();
  const isOpen = categoryDropdown.style.display === 'block';
  
  // Close all other dropdowns before opening this one
  if (!isOpen) {
    // Close all select dropdowns
    document.querySelectorAll('.select-dropdown').forEach(otherDropdown => {
      otherDropdown.style.display = 'none';
      const otherContainer = otherDropdown.closest('.select-container');
      if (otherContainer) {
        const otherArrow = otherContainer.querySelector('.select-arrow');
        if (otherArrow) {
          otherArrow.style.transform = 'rotate(0deg)';
        }
      }
    });
  }
  
  categoryDropdown.style.display = isOpen ? 'none' : 'block';
  categorySelectArrow.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(180deg)';
});

document.addEventListener('click', function(e) {
  if (!categorySelectContainer.contains(e.target)) {
    categoryDropdown.style.display = 'none';
    categorySelectArrow.style.transform = 'rotate(0deg)';
  }
});

categorySelectContainer.appendChild(categorySelectDisplay);
categorySelectContainer.appendChild(categoryDropdown);
categorySelectContainer.appendChild(categorySelect);

const itemContainer = document.createElement('div');
itemContainer.className = 'item-container';
itemContainer.className = 'item-container';

const itemTitleInput = document.createElement('input');
itemTitleInput.id = 'item-input';
itemTitleInput.name = 'item-input'; // Add name attribute to support autofill
itemTitleInput.type = 'text';
itemTitleInput.placeholder = '輸入項目名稱...';


itemContainer.appendChild(itemTitleInput);

const div2 = document.createElement('div');
div2.style.display = 'none';
const div3 = document.createElement('div');
div3.style.display = 'none';
const div4 = document.createElement('div');
div4.style.display = 'none';

const saveButton = document.createElement('button');
saveButton.textContent = '儲存';
saveButton.className = 'save-button';
// Initial state: disabled (gray)
saveButton.disabled = true;
saveButton.style.opacity = '0.5';
saveButton.style.cursor = 'not-allowed';

// Mark as changed, enable save button
const markAsChanged = () => {
  hasUnsavedChanges = true;
  if (saveButton) {
    saveButton.disabled = false;
    saveButton.style.opacity = '1';
    saveButton.style.cursor = 'pointer';
  }
};

// Mark as saved, disable save button
const markAsSaved = () => {
  hasUnsavedChanges = false;
  if (saveButton) {
    saveButton.disabled = true;
    saveButton.style.opacity = '0.5';
    saveButton.style.cursor = 'not-allowed';
  }
};

// Store current input values as original values (for comparison to detect changes)
const storeOriginalValues = () => {
  const itemInput = document.getElementById('item-input');
  const costInput = document.getElementById('cost-input');
  const noteInput = document.getElementById('note-input');
  const categorySelectElement = document.getElementById('expense-category-select');
  const categorySelect = document.getElementById('category-select');
  
  originalValues = {
    item: itemInput ? itemInput.value : '',
    cost: costInput ? costInput.value : '',
    note: noteInput ? noteInput.value : '',
    expenseCategory: categorySelectElement ? categorySelectElement.value : '',
    category: categorySelect ? categorySelect.value : ''
  };
};

// Check for changes, enable/disable save button based on result
const checkForChanges = () => {
  // Add new mode: any input is considered a change
  if (isNewMode) {
    const itemInput = document.getElementById('item-input');
    const costInput = document.getElementById('cost-input');
    const noteInput = document.getElementById('note-input');
    
    // Check if there's any valid input
    const hasInput = (itemInput && itemInput.value.trim() !== '') ||
                     (costInput && costInput.value.trim() !== '') ||
                     (noteInput && noteInput.value.trim() !== '');
    
    if (hasInput) {
      markAsChanged();
    } else {
      hasUnsavedChanges = false;
      if (saveButton) {
        saveButton.disabled = true;
        saveButton.style.opacity = '0.5';
        saveButton.style.cursor = 'not-allowed';
      }
    }
    return;
  }
  
  // Edit mode: compare current values with original values
  if (!originalValues) {
    return;
  }
  
  const itemInput = document.getElementById('item-input');
  const costInput = document.getElementById('cost-input');
  const noteInput = document.getElementById('note-input');
  const categorySelectElement = document.getElementById('expense-category-select');
  const categorySelect = document.getElementById('category-select');
  
  const currentItem = itemInput ? itemInput.value : '';
  const currentCost = costInput ? costInput.value : '';
  const currentNote = noteInput ? noteInput.value : '';
  const currentExpenseCategory = categorySelectElement ? categorySelectElement.value : '';
  const currentCategory = categorySelect ? categorySelect.value : '';
  
  // Compare if any field has changed
  const hasChanges = 
    currentItem !== originalValues.item ||
    currentCost !== originalValues.cost ||
    currentNote !== originalValues.note ||
    currentExpenseCategory !== originalValues.expenseCategory ||
    currentCategory !== originalValues.category;
  
  if (hasChanges) {
    markAsChanged();
  } else {
    hasUnsavedChanges = false;
    if (saveButton) {
      saveButton.disabled = true;
      saveButton.style.opacity = '0.5';
      saveButton.style.cursor = 'not-allowed';
    }
  }
};

const columnsContainer = document.createElement('div');
columnsContainer.className = 'columns-container';

const incomeColumn = document.createElement('div');
incomeColumn.className = 'income-column';

const incomeTitle = document.createElement('h3');
incomeTitle.className = 'income-title';
incomeTitle.textContent = '收入：';

const incomeAmount = document.createElement('div');
incomeAmount.className = 'income-amount';
incomeAmount.textContent = 0;

const expenseColumn = document.createElement('div');
expenseColumn.className = 'expense-column';

const expenseTitle = document.createElement('h3');
expenseTitle.className = 'expense-title';
expenseTitle.textContent = '預計支出：';

const expenseAmount = document.createElement('div');
expenseAmount.className = 'expense-amount';
expenseAmount.textContent = 0;

const totalColumn = document.createElement('div');
totalColumn.className = 'total-column';

const totalTitle = document.createElement('h3');
totalTitle.className = 'total-title';
totalTitle.textContent = '總計：';

const totalAmount = document.createElement('div');
totalAmount.className = 'total-amount';
totalAmount.textContent = 0;

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

// Month selection dropdown
let monthSelect = null;
const monthSelectWrapper = document.createElement('div');
monthSelectWrapper.className = 'month-select-wrapper';

const monthSelectLabel = document.createElement('span');
monthSelectLabel.className = 'month-select-label';
monthSelectLabel.textContent = '月份：';

monthSelect = document.createElement('select');
monthSelect.id = 'month-select';
monthSelect.name = 'month-select';
monthSelect.className = 'month-select';

monthSelectWrapper.appendChild(monthSelectLabel);
monthSelectWrapper.appendChild(monthSelect);

const submitContainer = document.createElement('div');
submitContainer.style.width = '100%';
submitContainer.style.display = 'flex';
submitContainer.style.justifyContent = 'center';
submitContainer.style.padding = '0';

// Load month names list
async function loadMonthNames() {
    const params = { name: "Show Tab Name" };
  const url = `${baseBudget}?${new URLSearchParams(params)}&_t=${Date.now()}`;
  const res = await fetch(url, {
      method: "GET",
      redirect: "follow",
    mode: "cors",
    cache: "no-store"
    });

    if (!res.ok) {
      throw new Error(`HTTP ${res.status}: ${res.statusText}`);
    }

    const data = await res.json();

    if (!Array.isArray(data) || data.length === 0) {
      return [];
    }

    sheetNames = data;
    return sheetNames;
}

// Initialize month dropdown
const initMonthSelect = async () => {
  if (!monthSelect) return;

  // First load month list (if not already loaded)
  if (sheetNames.length === 0) {
    try {
      await loadMonthNames();
      // Save to cache
      setToCache('budget_sheetNames', sheetNames);
    } catch (e) {
      // If loading fails, try reading from cache
      try {
        const storedSheetNames = await getFromCache('budget_sheetNames');
        if (storedSheetNames) {
          sheetNames = storedSheetNames;
        }
      } catch (e2) {
        // Cache may be unavailable or data corrupted, ignore error
      }
    }
  }

  try {
    // Directly infer current month and next month, don't wait for sheet name list
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const currentMonthStr = `${currentYear}${String(currentMonth).padStart(2, '0')}`;

    // Calculate next month
    let nextYear = currentYear;
    let nextMonth = currentMonth + 1;
    if (nextMonth > 12) {
      nextMonth = 1;
      nextYear = currentYear + 1;
    }
    const nextMonthStr = `${nextYear}${String(nextMonth).padStart(2, '0')}`;

    // Fix: directly find month index from sheetNames array, avoid errors from hardcoded reference points
    // Calculate next month's index
    const nextMonthArrayIndex = sheetNames.findIndex(name => name === nextMonthStr);
    const nextMonthIndex = nextMonthArrayIndex !== -1 ? nextMonthArrayIndex + 2 : -1;

    // Calculate current month's index
    const currentMonthArrayIndex = sheetNames.findIndex(name => name === currentMonthStr);
    const currentMonthIndex = currentMonthArrayIndex !== -1 ? currentMonthArrayIndex + 2 : -1;
    let targetSheetIndex = null;
    let targetMonthName = null;

    // First try loading all months' data from cache
    const monthsToCheck = [
      { index: nextMonthIndex, name: nextMonthStr },
      { index: currentMonthIndex, name: currentMonthStr }
    ];

    for (const { index, name } of monthsToCheck) {
      if (index < 2) continue;

      // First check if data exists in memory
      if (allMonthsData[index]) {
        const testData = allMonthsData[index];
        if (testData && testData.data) {
          const dataKeys = Object.keys(testData.data);
          const hasMonth = dataKeys.some(key => key.includes(name));
          if (hasMonth && !targetSheetIndex) {
            // If data has no explicit source marking, assume loaded from API (conservative approach)
            // This avoids incorrectly showing spinner
            if (testData._fromApi === undefined) {
              testData._fromApi = true;
            }
            targetSheetIndex = index;
            targetMonthName = name;
            break;
          }
        }
        // If has data but doesn't match conditions, continue checking next month
        continue;
      }

      // Try reading from cache
      try {
        const storedData = await getFromCache(`budget_monthData_${index}`);
        if (storedData) {
          const testData = storedData;
          if (testData && testData.data) {
            const dataKeys = Object.keys(testData.data);
            const hasMonth = dataKeys.some(key => key.includes(name));
            if (hasMonth) {
              // Explicitly mark as loaded from cache
              testData._fromCache = true;
              allMonthsData[index] = testData;
              if (!targetSheetIndex) {
                targetSheetIndex = index;
                targetMonthName = name;
                break; // Found target month, stop checking
              }
            }
          }
        }
      } catch (e) {
        // Cache may be unavailable or data corrupted, ignore error
      }
    }

    // If not found in cache, then send API request
    // Prefer trying next month (inferred index)
    if (!targetSheetIndex && nextMonthIndex >= 2) {
      try {
        const testData = await loadMonthData(nextMonthIndex);

        if (testData && testData.data) {
          // Verify returned data contains next month's named range
          const dataKeys = Object.keys(testData.data);
          const hasNextMonth = dataKeys.some(key => key.includes(nextMonthStr));

          if (hasNextMonth) {
            targetSheetIndex = nextMonthIndex;
            targetMonthName = nextMonthStr;
            testData._fromApi = true; // Mark as loaded from API
            allMonthsData[nextMonthIndex] = testData;

            // Save to cache
            setToCache(`budget_monthData_${nextMonthIndex}`, testData);
          }
        }
      } catch (e) {
      }
    }

    // If next month loading failed, try current month
    if (!targetSheetIndex && currentMonthIndex >= 2) {
      try {
        const testData = await loadMonthData(currentMonthIndex);

        if (testData && testData.data) {
          const dataKeys = Object.keys(testData.data);
          const hasCurrentMonth = dataKeys.some(key => key.includes(currentMonthStr));
          if (hasCurrentMonth) {
            targetSheetIndex = currentMonthIndex;
            targetMonthName = currentMonthStr;
            testData._fromApi = true; // Mark as loaded from API
            allMonthsData[currentMonthIndex] = testData;

            // Save to cache
            setToCache(`budget_monthData_${currentMonthIndex}`, testData);
          }
        }
      } catch (e) {
      }
    }

    // If found target month, display it
    // If target month not found but current month is in sheetNames, try using current month
    if (!targetSheetIndex && sheetNames.length > 0) {
      const currentMonthInSheetNames = sheetNames.findIndex(name => name === currentMonthStr);
      if (currentMonthInSheetNames >= 0) {
        const currentMonthSheetIndex = currentMonthInSheetNames + 2;
        if (allMonthsData[currentMonthSheetIndex]) {
          targetSheetIndex = currentMonthSheetIndex;
          targetMonthName = currentMonthStr;
        } else {
          // If current month not in memory, try loading
          try {
            const testData = await loadMonthData(currentMonthSheetIndex);
            if (testData && testData.data) {
              const dataKeys = Object.keys(testData.data);
              const hasCurrentMonth = dataKeys.some(key => key.includes(currentMonthStr));
              if (hasCurrentMonth) {
                targetSheetIndex = currentMonthSheetIndex;
                targetMonthName = currentMonthStr;
                testData._fromApi = true;
                allMonthsData[currentMonthSheetIndex] = testData;
                setToCache(`budget_monthData_${currentMonthSheetIndex}`, testData);
              }
            }
          } catch (e) {
          }
        }
      }
    }
    
    if (targetSheetIndex && targetMonthName) {
      currentSheetIndex = targetSheetIndex;
      const targetMonthData = allMonthsData[targetSheetIndex];
      // Check if data loaded from cache (needs to cover header)
      // Only consider loaded from cache when _fromCache is true or _fromApi is explicitly false
      const isFromCache = allMonthsData[targetSheetIndex] &&
                          (targetMonthData._fromCache === true || targetMonthData._fromApi === false);

      // If loaded from cache, first hide original spinner, then show spinner covering header
      if (isFromCache) {
        hideSpinner();
        showSpinner(true); // Cover header
      }

      // Immediately process and display target month data
      if (targetMonthData.data) {
        processDataFromResponse(targetMonthData.data, true);

        // Update total immediately after processing data (ensure total auto-calculated)
        if (targetMonthData.total && Array.isArray(targetMonthData.total) && targetMonthData.total.length >= 3) {
          updateTotalDisplay(targetMonthData.total);
        } else {
          // If no total or format incorrect, reload total
          try {
            await loadMonthData(targetSheetIndex, false);
            // loadMonthData will automatically update total
          } catch (error) {
            // If loading fails, use live calculation
            updateTotalDisplay();
          }
        }
      } else {
        // If no data, reload
        try {
          await loadMonthData(targetSheetIndex, false);
        } catch (error) {
          allRecords = [];
          filteredRecords = [];
          // Even if no data, update total (display as 0)
          updateTotalDisplay();
        }
      }

      // Ensure first record displayed (if any, and not in add new mode)
      if (filteredRecords.length > 0) {
        showRecord(0);
        updateArrowButtons();
      } else {
        // If no records, enter add new mode
      enterNewModeIfEmpty();
      }
      updateDeleteButton();

      // If loaded from cache, delay hiding spinner (let user see loading process)
      if (isFromCache) {
        setTimeout(() => {
      hideSpinner();
        }, 100);
      } else {
        // Loaded from API, immediately close loading overlay (don't cover header)
        hideSpinner();
      }

      // Directly fill dropdown using inference (limit to at most next month)
      monthSelect.innerHTML = '';
      const monthOptions = [];

      // Add current month
      monthOptions.push({ name: currentMonthStr, index: currentMonthIndex });
      // Add next month (if different)
      if (nextMonthStr !== currentMonthStr) {
        monthOptions.push({ name: nextMonthStr, index: nextMonthIndex });
      }

      // Create options (sorted by time)
      monthOptions.sort((a, b) => {
        const yearA = parseInt(a.name.substring(0, 4));
        const monthA = parseInt(a.name.substring(4, 6));
        const yearB = parseInt(b.name.substring(0, 4));
        const monthB = parseInt(b.name.substring(4, 6));
        if (yearA !== yearB) return yearA - yearB;
        return monthA - monthB;
      });

      monthOptions.forEach((option, idx) => {
        const opt = document.createElement('option');
        opt.value = String(option.index - 2); // Option value is sheetIndex - 2 (corresponds to sheetNames index)
        opt.textContent = option.name;
        monthSelect.appendChild(opt);
      });

      // Set currently selected month
      if (targetSheetIndex >= 2) {
        const selectIndex = targetSheetIndex - 2;
        const foundOption = monthOptions.findIndex(opt => opt.index === targetSheetIndex);
        if (foundOption >= 0) {
          monthSelect.value = String(monthOptions[foundOption].index - 2);
        }
      }

      // Load full month list in background and preload other months' data
      loadMonthNames().then(() => {
        // Update dropdown with complete month list
        monthSelect.innerHTML = '';
        sheetNames.forEach((name, idx) => {
          const opt = document.createElement('option');
          opt.value = String(idx);
          opt.textContent = name;
          monthSelect.appendChild(opt);
        });
        // Update month selector display
        if (currentSheetIndex >= 2) {
          const selectIndex = currentSheetIndex - 2;
          if (selectIndex >= 0 && selectIndex < sheetNames.length) {
            monthSelect.value = String(selectIndex);
          }
        }

        // Start preloading other months' data in background
        preloadAllMonthsData(0, 0).then(() => {
        }).catch((error) => {
          // Background preload failure doesn't affect user operation
        });
      }).catch((err) => {
        // Month list loading failure doesn't affect current display
        // Even if failed, try showing loaded months (if any)
        if (sheetNames.length > 0) {
          monthSelect.innerHTML = '';
          sheetNames.forEach((name, idx) => {
            const opt = document.createElement('option');
            opt.value = String(idx);
            opt.textContent = name;
            monthSelect.appendChild(opt);
          });
        }
      });
    } else {
      // If nothing found, try reading from memory
      hideSpinner();
      if (await loadContentFromMemory()) {
        enterNewModeIfEmpty();
        updateDeleteButton();
        updateArrowButtons();
      }
    }

    // When dropdown changes, switch month (read from memory, no request sent)
    // First remove old event listener (if exists)
    if (monthSelectChangeHandler) {
      monthSelect.removeEventListener('change', monthSelectChangeHandler);
    }

    monthSelectChangeHandler = async () => {
      // Prevent rapid consecutive switching
      if (isSwitchingMonth) {
        return;
      }

      const idx = parseInt(monthSelect.value, 10);
      if (!Number.isFinite(idx) || idx < 0 || idx >= sheetNames.length) {
        return;
      }

      // If selecting current month, no need to switch
      const newSheetIndex = idx + 2;
      const monthName = sheetNames[idx];

      if (currentSheetIndex === newSheetIndex) {
        return;
      }

      isSwitchingMonth = true;
      const oldSheetIndex = currentSheetIndex;
      currentSheetIndex = newSheetIndex; // Convert to actual sheet index (first two are "blank sheet" and "dropdown")

      // When switching months, leave add new mode and reset state
      isNewMode = false;
      currentRecordNumber = null;

          const itemInput = document.getElementById('item-input');
          const costInput = document.getElementById('cost-input');
          const noteInput = document.getElementById('note-input');
      const expenseCategorySelect = document.getElementById('expense-category-select');

      try {
        // First try loading data from memory (no request, no loading animation)
        if (await loadContentFromMemory()) {
          // Loaded from memory successfully, immediately update UI
          updateDeleteButton();
          updateArrowButtons();
          enterNewModeIfEmpty();
          isSwitchingMonth = false;
        } else {
          // When no data in memory, show progress bar and load
          showSpinner();

          try {
            // Load that month's data
            const monthData = await loadMonthData(currentSheetIndex);
            allMonthsData[currentSheetIndex] = monthData;

            // Save to cache
            setToCache(`budget_monthData_${currentSheetIndex}`, monthData);

            // Process and display data
            processDataFromResponse(monthData.data);
            updateTotalDisplay(monthData.total);
            updateDeleteButton();
            updateArrowButtons();
            enterNewModeIfEmpty();

            // Close progress bar after loading complete
            hideSpinner();
          } catch (error) {
            alert(`載入月份 ${monthName} 失敗: ${error.message || error.toString()}`);
            // Restore to original month
            currentSheetIndex = oldSheetIndex;
            if (monthSelect) {
              const oldSelectIndex = oldSheetIndex - 2;
              if (oldSelectIndex >= 0 && oldSelectIndex < sheetNames.length) {
                monthSelect.value = String(oldSelectIndex);
              }
            }
            hideSpinner();
          } finally {
            isSwitchingMonth = false;
          }
        }
      } catch (e) {
        // If loading fails, restore original month
        currentSheetIndex = oldSheetIndex;
        if (monthSelect) {
          const oldSelectIndex = oldSheetIndex - 2;
          if (oldSelectIndex >= 0 && oldSelectIndex < sheetNames.length) {
            monthSelect.value = String(oldSelectIndex);
          }
        }
        isSwitchingMonth = false;
          hideSpinner();
      }
    };

    monthSelect.addEventListener('change', monthSelectChangeHandler);

    // Initially load currently selected month data (read from memory, no spinner needed)
    // If no data in memory, initMonthSelect has already loaded and displayed
    if (await loadContentFromMemory()) {
      // Loaded from memory successfully, only need to update UI
    updateDeleteButton();
    updateArrowButtons();
    }
  } catch (error) {
    // CORS / fetch type errors don't affect results in actual use, no alert popup here to avoid interfering with operations
  }
};


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
budgetCardsContainer.appendChild(recordNumber);
budgetCardsContainer.appendChild(headerInfo);
budgetCardsContainer.appendChild(deleteButton);
budgetCardsContainer.appendChild(leftArrow);
budgetCardsContainer.appendChild(rightArrow);
budgetCardsContainer.appendChild(itemContainer);
budgetCardsContainer.appendChild(div1);
  div1.appendChild(categoryLabel);
  div1.appendChild(categorySelectContainer);
budgetCardsContainer.appendChild(div2);
budgetCardsContainer.appendChild(div3);
budgetCardsContainer.appendChild(div4);
budgetCardsContainer.appendChild(submitContainer);
  submitContainer.appendChild(saveButton);

categorySelect.addEventListener('change', () => {
  updateDivVisibility();
  if (allRecords.length > 0) {
    filterRecordsByType(categorySelect.value);
  }
});
updateDivVisibility();
saveButton.addEventListener('click', saveData);
saveButton.addEventListener('click', loadTotal);

// Add bounce prevention for dropdown container (main page handled by CSS)
// (function() {
//   // Check if is drag element
//   function isDragElement(target) {
//     return target.closest('.drag-handle') ||
//            target.closest('.option-item') ||
//            target.closest('[draggable="true"]') ||
//            target.classList.contains('drag-handle') ||
//            target.classList.contains('option-item') ||
//            target.hasAttribute('draggable');
//   }

//   // Hard lock solution: lock body scroll when dropdown opens (prevent entire page from being pulled)
//   let bodyScrollLocked = false;
//   let scrollY = 0;

//   function lockBodyScroll() {
//     if (!bodyScrollLocked) {
//       // Remember current scroll position
//       scrollY = window.scrollY || window.pageYOffset || document.documentElement.scrollTop;

//       // Lock both html and body (on iOS the actual scroll container is sometimes html)
//       document.documentElement.style.overflow = 'hidden';
//       document.body.style.overflow = 'hidden';

//       // Use fixed positioning to maintain visual position, avoid jumping
//       document.body.style.position = 'fixed';
//       document.body.style.top = `-${scrollY}px`;
//       document.body.style.width = '100%';

//       bodyScrollLocked = true;
//     } else {
//     }
//   }

//   function unlockBodyScroll() {
//     if (bodyScrollLocked) {

//       // Restore html and body overflow
//       document.documentElement.style.overflow = '';
//       document.body.style.overflow = '';

//       // Restore body positioning styles
//       document.body.style.position = '';
//       document.body.style.top = '';
//       document.body.style.width = '';

//       // Restore to original scroll position
//       window.scrollTo(0, scrollY);

//       bodyScrollLocked = false;
//     } else {
//     }
//   }

//   // Listen to all dropdown show/hide
//   let dropdownObserver = null;
//   let styleObserver = null;
//   const observedDropdowns = new Set();

//   function checkDropdownsAndLock() {
//     // Re-query all dropdowns (because they're dynamically created)
//     const dropdowns = document.querySelectorAll('.category-dropdown, .select-dropdown');

//     let anyOpen = false;
//     const openDropdowns = [];

//     dropdowns.forEach((dropdown, index) => {
//       const isOpen = dropdown.style.display === 'block' ||
//                     window.getComputedStyle(dropdown).display === 'block';
//       if (isOpen) {
//         anyOpen = true;
//         openDropdowns.push(index);
//       }

//       // Set up listeners for newly created dropdowns
//       if (!observedDropdowns.has(dropdown)) {
//         observedDropdowns.add(dropdown);
//         if (styleObserver) {
//           styleObserver.observe(dropdown, {
//             attributes: true,
//             attributeFilter: ['style']
//           });
//         }
//       }
//     });

//     if (openDropdowns.length > 0) {
//     }

//     if (anyOpen) {
//       lockBodyScroll();
//     } else {
//       unlockBodyScroll();
//     }
//   }

//   function setupBodyScrollLock() {

//     // Create observer to listen for dropdown style changes
//     styleObserver = new MutationObserver(() => {
//       checkDropdownsAndLock();
//     });

//     // Create observer to listen for DOM changes (when new dropdown is created)
//     dropdownObserver = new MutationObserver((mutations) => {
//       let hasNewDropdown = false;
//       mutations.forEach(mutation => {
//         mutation.addedNodes.forEach(node => {
//           if (node.nodeType === 1) { // Element node
//             if (node.classList && (
//               node.classList.contains('category-dropdown') ||
//               node.classList.contains('select-dropdown')
//             )) {
//               hasNewDropdown = true;
//             }
//             // Check child nodes
//             if (node.querySelectorAll) {
//               const childDropdowns = node.querySelectorAll('.category-dropdown, .select-dropdown');
//               if (childDropdowns.length > 0) {
//                 hasNewDropdown = true;
//               }
//             }
//           }
//         });
//       });

//       if (hasNewDropdown) {
//         checkDropdownsAndLock();
//       }
//     });

//     // Listen to entire body changes
//     dropdownObserver.observe(document.body, {
//       childList: true,
//       subtree: true
//     });

//     // Initial check once
//     checkDropdownsAndLock();

//   }

//   // Add bounce prevention for all dropdown containers
//   function setupDropdownPrevention() {
//     const dropdowns = document.querySelectorAll('.category-dropdown, .select-dropdown, .options-list, .select-options');
//     dropdowns.forEach(dropdown => {
//       // Avoid duplicate binding
//       if (dropdown.dataset.bouncePrevented) return;
//       dropdown.dataset.bouncePrevented = 'true';

//       let dropdownTouchStartY = 0;

//       dropdown.addEventListener('touchstart', function(e) {
//         if (isDragElement(e.target) || window._isDragging) {
//           return;
//         }
//         dropdownTouchStartY = e.touches[0].clientY;
//       }, { passive: true });

//       dropdown.addEventListener('touchmove', function(e) {
//         if (window._isDragging || isDragElement(e.target)) {
//           return;
//         }

//         const currentY = e.touches[0].clientY;
//         const deltaY = currentY - dropdownTouchStartY;
//         const currentScrollTop = dropdown.scrollTop;
//         const scrollHeight = dropdown.scrollHeight;
//         const clientHeight = dropdown.clientHeight;

//         // Precise judgment: only prevent when at boundary and continuing to slide toward boundary
//         // Use <= 1 instead of === 0 to handle negative values and sub-pixel rounding (consistent with settings page)
//         const isAtTop = currentScrollTop <= 1;
//         const isAtBottom = currentScrollTop + clientHeight >= scrollHeight - 1;

//         // Prevent when at top and pulling down (deltaY > 0)
//         if (isAtTop && deltaY > 0) {
//           if (e.cancelable) {
//             e.preventDefault();
//           }
//           dropdown.scrollTop = 0;
//         }
//         // Prevent when at bottom and pulling up (deltaY < 0)
//         else if (isAtBottom && deltaY < 0) {
//           if (e.cancelable) {
//             e.preventDefault();
//           }
//           dropdown.scrollTop = Math.max(0, scrollHeight - clientHeight);
//         }
//       }, { passive: false });
//     });
//   }

//   // Set up dropdowns after page loads
//   if (document.readyState === 'loading') {
//     document.addEventListener('DOMContentLoaded', () => {
//       setupDropdownPrevention();
//       // Enable hard lock solution (prevent entire page from being pulled)
//       setupBodyScrollLock();
//     });
//   } else {
//     setupDropdownPrevention();
//     // Enable hard lock solution (prevent entire page from being pulled)
//     setupBodyScrollLock();
//   }

//   // Listen to dynamically added dropdowns
//   const observer = new MutationObserver(setupDropdownPrevention);
//   observer.observe(document.body, { childList: true, subtree: true });
// })();

document.addEventListener('DOMContentLoaded', async function() {
  // Clear old cache data, force reload from API
  // Clear data in memory
  Object.keys(allMonthsData).forEach(key => delete allMonthsData[key]);

  // Immediately show loading animation to let user know page is loading
  showSpinner();

  // Send Create Tab request as soon as page loads
  try {
    await callAPI({ name: "Create Tab" });
  } catch (e) {
    // Creation failed, ignore error (may already exist)
  }

  document.getElementsByClassName('post-content')[0].appendChild(totalContainer);
  document.getElementsByClassName('post-content')[0].appendChild(budgetCardsContainer);

  try {
    // Non-blocking load latest options from "dropdown" sheet
    const refreshDropdowns = () => {
      loadDropdownOptions().then(() => {
        // If currently showing expense category, re-render
        const categorySelect = document.getElementById('category-select');
        if (categorySelect && categorySelect.value === '預計支出') {
          updateDivVisibility('預計支出');
        }
      }).catch(err => {
      });
    };

    refreshDropdowns();

    // Listen for settings page update notifications (auto-reload after settings page updates dropdown)
    window.addEventListener('storage', (e) => {
      if (e.key === 'dropdownUpdated') {
        refreshDropdowns();
      }
    });

    // Listen for manual sync requests (triggered from sync icon in navigation bar)
    window.addEventListener('syncRequested', async () => {
      try {
        SyncStatus.startSync();
        // Clear memory cache and reload current month
        delete allMonthsData[currentSheetIndex];
        await loadContent(true); // Force reload
        SyncStatus.endSync(true);
      } catch (e) {
        SyncStatus.endSync(false);
      }
    });

    // Also listen for storage events on the same page (because storage events only fire on other tabs)
    let lastUpdateTime = localStorage.getItem('dropdownUpdated');
    setInterval(() => {
      const current = localStorage.getItem('dropdownUpdated');
      if (current && current !== lastUpdateTime) {
        lastUpdateTime = current;
        refreshDropdowns();
      }
    }, 1000); // Check every second

    // initMonthSelect internally handles loading and hides spinner when complete
    await initMonthSelect(); // First load month list and corresponding month data
    updateDeleteButton(); // Initialize delete button display state
  } catch (error) {
    hideSpinner(); // Ensure spinner is hidden on error
    const errorContainer = document.createElement('div');
    errorContainer.innerHTML = '載入失敗: ' + error.message;
    errorContainer.style.color = 'red';
    errorContainer.style.marginTop = '20px';
    document.getElementsByClassName('post-content')[0].appendChild(errorContainer);
  }
});

</script>