function doGet(e) {
  var name = e.parameter.name;
  var sheet = e.parameter.sheet;
  if (name === "Show Tab Name") {
    var tab = ShowTabName();
    return _json(tab);
  } else if (name === "Show Tab Data") {
    var data = ShowTabData(sheet);
    return _json(data);
  } else if (name === "Show Total"){
    var data = GetSummary(sheet)
    return _json(data)
  }
}

function doPost(e) {
  var contents = JSON.parse(e.postData.contents)
  var name = contents.name;
  var sheet = contents.sheet;
  var range = contents.range;
  var sheetName = contents.sheetName;
  var category = contents.category;
  var item = contents.item;
  var cost = contents.cost;
  var note = contents.note;
  var number = contents.number
  var updateRow = contents.updateRow;
  var oldValue = contents.oldValue
  var newValue = contents.newValue
  var itemId = contents.itemId
  var action = contents.action
  // New: receive drag-and-drop sort indices
  var oldIndex = contents.oldIndex;
  var newIndex = contents.newIndex;
  // New: receive batch update data
  var originalData = contents.originalData;
  var newData = contents.newData;

  var result = { success: false, message: "" };

  try {
    if (name === "Upsert Data") {
      result = UpsertData(sheet,range,category,item,cost,note,updateRow);
    } else if (name === "Create Tab") {
      result = CreateNewTab();
    } else if (name === "Delete Data") {
      result = DeleteData(sheet,range, number);
    } else if (name === "Delete Tab") {
      result = DeleteTab(sheet);
    } else if (name === "Change Tab Name") {
      result = ChangeTabName(sheet,sheetName);
    } else if (name === "Update Dropdown"){
      result = UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex);
    } else if (name === "Batch Update Dropdown") {
      result = BatchUpdateDropdown(itemId, originalData, newData);
    } else {
      result.message = "未知的操作類型: " + e.parameter;
    }
  } catch (error) {
    result.success = false;
    result.message = "操作失敗: " + error.toString();
  }

  return _json(result);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cache helper functions
var CACHE_EXPIRATION = 300; // 5 minutes

function getCacheKey(prefix, sheetIndex) {
  return prefix + '_' + (sheetIndex || 'all');
}

function getFromCache(key) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      return null;
    }
  }
  return null;
}

function setToCache(key, data) {
  var cache = CacheService.getScriptCache();
  try {
    var jsonStr = JSON.stringify(data);
    // Only cache if data is less than 100KB (CacheService limit)
    if (jsonStr.length < 100000) {
      cache.put(key, jsonStr, CACHE_EXPIRATION);
    }
  } catch (e) {
    // Ignore cache errors
  }
}

function invalidateCache(sheetIndex) {
  var cache = CacheService.getScriptCache();
  cache.remove(getCacheKey('tabData', sheetIndex));
  cache.remove(getCacheKey('summary', sheetIndex));
}

// Clear month list cache (used when adding new tab)
function invalidateTabNamesCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(getCacheKey('tabNames'));
}

function ShowTabName() {
  var cacheKey = getCacheKey('tabNames');
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var sheetNames = [];
  
  // Start from index 2 (skip first two: "blank sheet", "dropdown")
  // slice(2) will include all elements from index 2 to the end
  var sheetsToProcess = allSheets.slice(2);
  sheetsToProcess.forEach(function(sheet){
    sheetNames.push(sheet.getSheetName());
  });

  setToCache(cacheKey, sheetNames);
  return sheetNames;
}

// Batch update dropdown
function BatchUpdateDropdown(itemId, originalData, newData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dropdownSheet = ss.getSheets()[1];
  var data = dropdownSheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, message: '下拉選單表是空的' };
  }

  var headerRow = data[0];
  var colIndex = -1;

  // Find corresponding column
  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['支出－項目', '支出-項目', '消費類別', '類別']);
  } else if (itemId === 'payment') {
    colIndex = FindHeaderColumn(headerRow, ['支付方式']);
  } else if (itemId === 'platform') {
    colIndex = FindHeaderColumn(headerRow, ['支付平台', '平台']);
  }

  if (colIndex < 0) {
    return { success: false, message: '找不到對應欄位' };
  }

  // Validate data
  if (!Array.isArray(originalData) || !Array.isArray(newData)) {
    return { success: false, message: 'originalData 和 newData 必須是陣列' };
  }

  // Find deleted items (in original data but not in new data)
  var removedItems = originalData.filter(function(item) {
    return newData.indexOf(item) === -1;
  });

  // Find added items (in new data but not in original data)
  var addedItems = newData.filter(function(item) {
    return originalData.indexOf(item) === -1;
  });

  // Detect rename: if deleted and added counts are the same, assume it's a rename
  var renames = [];
  if (removedItems.length === addedItems.length && removedItems.length > 0) {
    for (var i = 0; i < removedItems.length; i++) {
      renames.push({
        oldValue: removedItems[i],
        newValue: addedItems[i]
      });
    }
  }

  // 1. Update dropdown table: clear the column and write new data
  var maxRows = dropdownSheet.getMaxRows();
  if (maxRows > 1) {
    dropdownSheet.getRange(2, colIndex + 1, maxRows - 1, 1).clearContent();
  }

  // Write new data
  if (newData.length > 0) {
    var writeData = newData.map(function(item) { return [item]; });
    dropdownSheet.getRange(2, colIndex + 1, newData.length, 1).setValues(writeData);
  }

  // 2. If it's category, update historical data in all monthly sheets
  var historicalUpdates = 0;
  if (itemId === 'category' && renames.length > 0) {
    var allSheets = ss.getSheets();
    var categoryColumnIndex = 8; // Column I (category)
    var expenseStartColumnIndex = 6; // Column G (expense record start column, number)

    for (var s = 2; s < allSheets.length; s++) {
      var monthSheet = allSheets[s];
      var monthData = monthSheet.getDataRange().getValues();

      if (monthData.length < 2) continue;

      for (var row = 1; row < monthData.length; row++) {
        if (monthData[row][0] === '總計' || monthData[row][0] === '') {
          continue;
        }

        var expenseNumber = monthData[row][expenseStartColumnIndex];
        if (expenseNumber !== '' && expenseNumber !== null && expenseNumber !== undefined) {
          var categoryValue = monthData[row][categoryColumnIndex];

          for (var r = 0; r < renames.length; r++) {
            if (categoryValue === renames[r].oldValue) {
              monthSheet.getRange(row + 1, categoryColumnIndex + 1).setValue(renames[r].newValue);
              historicalUpdates++;
              break;
            }
          }
        }
      }
    }
  }

  // Assemble return message
  var message = '下拉選單已更新';
  if (renames.length > 0) {
    message += '，重新命名 ' + renames.length + ' 項';
    if (historicalUpdates > 0) {
      message += '，歷史記錄更新 ' + historicalUpdates + ' 筆';
    }
  }
  if (removedItems.length > addedItems.length) {
    message += '，刪除 ' + (removedItems.length - addedItems.length) + ' 項';
  }
  if (addedItems.length > removedItems.length) {
    message += '，新增 ' + (addedItems.length - removedItems.length) + ' 項';
  }

  return {
    success: true,
    message: message,
    renames: renames,
    historicalUpdates: historicalUpdates
  };
}

function UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dropdownSheet = ss.getSheets()[1];
  var data = dropdownSheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, message: '下拉選單表是空的' };
  }

  var headerRow = data[0];
  var colIndex = -1;

  // In the dropdown table, category corresponds to column name "支出－項目"
  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['支出－項目', '支出-項目', '消費類別', '類別']);
  } else if (itemId === 'payment') {
    colIndex = FindHeaderColumn(headerRow, ['支付方式']);
  } else if (itemId === 'platform') {
    colIndex = FindHeaderColumn(headerRow, ['支付平台', '平台']);
  }

  if (colIndex < 0) {
    return { success: false, message: '找不到對應欄位' };
  }

  // === New: drag-and-drop sort feature ===
  if (action === 'reorder') {
    // Check index parameters
    if (typeof oldIndex !== 'number' || typeof newIndex !== 'number') {
      return { success: false, message: '排序時 oldIndex 和 newIndex 必須是數字' };
    }

    // Collect all values in this column (skip header row)
    var values = [];
    for (var r = 1; r < data.length; r++) {
      var val = data[r][colIndex];
      if (val !== '' && val !== null && val !== undefined) {
        values.push(val);
      }
    }

    // Check index range
    if (oldIndex < 0 || oldIndex >= values.length || newIndex < 0 || newIndex >= values.length) {
      return { success: false, message: '索引超出範圍' };
    }

    // Execute sort: move item from oldIndex position to newIndex position
    var movedItem = values.splice(oldIndex, 1)[0];
    values.splice(newIndex, 0, movedItem);

    // Update dropdown table (write starting from row 2)
    var writeRow = 2;
    for (var i = 0; i < values.length; i++) {
      dropdownSheet.getRange(writeRow, colIndex + 1).setValue(values[i]);
      writeRow++;
    }

    // Clear extra cells after
    if (writeRow <= data.length) {
      dropdownSheet.getRange(writeRow, colIndex + 1, data.length - writeRow + 1, 1).clearContent();
    }

    return { success: true, message: '排序成功，已更新下拉選單順序' };
  }

  var updatedCount = 0;

  // === New: only update dropdown table ===
  if (action === 'add') {
    if (!newValue) {
      return { success: false, message: '新增時 newValue 必填' };
    }

    // Find the last row with a value in this column (search from bottom up)
    var lastRowInColumn = 1; // Default to header row (index 1, corresponds to row 2)
    for (var r = data.length - 1; r >= 1; r--) { // Search from last row upward, skip header row
      var val = data[r][colIndex];
      if (val !== '' && val !== null && val !== undefined) {
        lastRowInColumn = r + 1; // Found last row with value (+1 to convert to row number)
        break;
      }
    }

    // Add at the next position after the last row in this column
    dropdownSheet.getRange(lastRowInColumn + 1, colIndex + 1).setValue(newValue);
    return { success: true, message: '已新增到下拉選單', updated: 1 };
  }

  // === Delete: only update dropdown table ===
  if (action === 'delete') {
    if (!oldValue) {
      return { success: false, message: '刪除時 oldValue 必填（不能是空白）' };
    }
    var writeRow = 2;
    for (var r = 1; r < data.length; r++) {
      var val = data[r][colIndex];
      if (val === oldValue) {
        updatedCount++;
        continue;
      }
      dropdownSheet.getRange(writeRow, colIndex + 1).setValue(val);
      writeRow++;
    }
    if (writeRow <= data.length) {
      dropdownSheet.getRange(writeRow, colIndex + 1, data.length - writeRow + 1, 1).clearContent();
    }
    return { success: true, message: '已從下拉選單刪除', updated: updatedCount };
  }

  // === Edit (merge): update dropdown table + historical data in all monthly sheets ===
  if (action === 'edit') {
    if (!oldValue || !newValue || oldValue === newValue) {
      return { success: false, message: '編輯時 oldValue / newValue 不正確（oldValue 不可空白）' };
    }

    // 1. Update dropdown table (column name: 支出－項目)
    for (var r2 = 1; r2 < data.length; r2++) {
      var val2 = data[r2][colIndex];
      if (val2 === oldValue) {
        data[r2][colIndex] = newValue;
        updatedCount++;
      }
    }
    dropdownSheet.getDataRange().setValues(data);

    // 2. Update historical data in all monthly sheets (only update "category" column for expense records with range = 0)
    // Only update monthly sheets when itemId === 'category'
    if (itemId === 'category') {
      var allSheets = ss.getSheets();
      var monthSheetsUpdated = 0;
      var totalRecordsUpdated = 0;

      // In budget table, expense records start from column G (index 6)
      // Column G (index 6): number
      // Column H (index 7): time
      // Column I (index 8): category - this is the column we want to update
      var categoryColumnIndex = 8; // Column I (category)
      var expenseStartColumnIndex = 6;  // Column G (expense record start column, number)

      for (var s = 2; s < allSheets.length; s++) { // Start from index 2 (skip first two sheets)
        var monthSheet = allSheets[s];
        var monthData = monthSheet.getDataRange().getValues();

        if (monthData.length < 2) continue; // Skip empty sheets

        var monthUpdated = false;

        // Start from row 2 (index 1), skip header row
        for (var row = 1; row < monthData.length; row++) {
          // Check if it's "總計" row or empty row
          if (monthData[row][0] === '總計' || monthData[row][0] === '') {
            continue;
          }

          // Only process expense records with range = 0 (records starting from column G)
          // Check if column G has a value (number), if so it's an expense record
          var expenseNumber = monthData[row][expenseStartColumnIndex]; // Column G (number)

          // Only check column I category if column G has a value (is expense record)
          if (expenseNumber !== '' && expenseNumber !== null && expenseNumber !== undefined) {
            var categoryValue = monthData[row][categoryColumnIndex]; // Column I (category)

            if (categoryValue === oldValue) {
              monthSheet.getRange(row + 1, categoryColumnIndex + 1).setValue(newValue);
              monthUpdated = true;
              totalRecordsUpdated++;
            }
          }
        }

        if (monthUpdated) {
          monthSheetsUpdated++;
        }
      }

      return {
        success: true,
        message: '已合併完成，下拉選單更新：' + updatedCount + ' 筆，月份表格更新：' + totalRecordsUpdated + ' 筆（' + monthSheetsUpdated + ' 個月份）'
      };
    } else {
      // If it's not category, only update dropdown table
      return {
        success: true,
        message: '已更新下拉選單：' + updatedCount + ' 筆'
      };
    }
  }

  return { success: false, message: '未知的 action: ' + action };
}

function ShowTabData(sheet) {
  var cacheKey = getCacheKey('tabData', sheet);
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var allRanges = ss.getNamedRanges();  // Get ALL named ranges from spreadsheet
  var res = {};
  allRanges.forEach(function(namedRange) {
    var rng = namedRange.getRange();
    // Only include ranges that belong to the target sheet
    if (rng.getSheet().getSheetId() === targetSheet.getSheetId()) {
      var title = namedRange.getName();
      res[title] = cleanValues(rng.getValues());
    }
  });

  function isEmpty(v) {
    return v === "" || v === undefined || v === null;
  }

  function cleanValues(values) {
    var res = [];
    values.forEach(function(value){
      // Keep the row if any field is not empty
      // This ensures even rows with only a number are preserved
      if(!value.every(isEmpty)) {
        res.push(value);
      }
    });
    return res;
  }

  setToCache(cacheKey, res);
  return res;
}

function FindHeaderColumn(headerRow, keywords) {
  for (var c = 0; c < headerRow.length; c++) {
    var headerText = (headerRow[c] || '').toString().trim();
    if (!headerText) continue;
    for (var k = 0; k < keywords.length; k++) {
      if (headerText.indexOf(keywords[k]) !== -1) {
        return c;
      }
    }
  }
  return -1;
}

function CreateNewTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var targetSheet = ss.getSheets()[0];
  var incomeValue = [["總計",,,"0"]]
  targetSheet.getRange("A3:D3").setValues(incomeValue)
  var expenseValue = [["總計",,,,"0"]]
  targetSheet.getRange("G3:K3").setValues(expenseValue)
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth()+1;
  var sheetNames = [];
  allSheets.forEach(function(sheet){
    sheetNames.push(sheet.getSheetName());
  });
  if (month < 10) {
    month = '0' + month;
  }
  if (sheetNames.includes(year.toString() + month.toString()) && parseInt(month)+1 <= 12) {
    month = parseInt(month) + 1;
  } else if (sheetNames.includes(year.toString() + month.toString()) && parseInt(month)+1 > 12){
    year = parseInt(year) + 1;
    month = 1;
  }
  if (parseInt(month) < 10) {
    month = '0' + parseInt(month);
  }
  if (sheetNames.includes(year.toString() + month.toString())) {
    return { success: false, message: 'Next month spreadsheet already created, please do not add again' };
  } else {
    var destination = SpreadsheetApp.openById('1P5RtR3fYgSVfvCjY5ryLqEfYU4jEvsacNsSg_1LAdQg');
    targetSheet.copyTo(destination);
    var copyName = ss.getSheetByName('「空白表」的副本');
    copyName.setName(year + month.toString());
  }
  month = month.toString();
  var targetSheet = ss.getSheetByName(year.toString()+month.toString());
  var range = targetSheet.getRange('A2:E');
  var name = '當月收入' + year.toString() + month.toString();
  ss.setNamedRange(name,range);
  var range = targetSheet.getRange('G2:L');
  var name = '當月支出預算' + year.toString() + month.toString();
  ss.setNamedRange(name,range);
  // Clear month list cache to ensure latest list is retrieved on next load
  invalidateTabNamesCache();
  return { success: true, message: 'New tab successfully created: ' + year + month };
}


function UpsertData(sheetIndex, rangeType, category, item, cost, note, updateRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 3; // 從第三行開始（Row1: 類型標題, Row2: 欄位標題, Row3: 資料開始）
  var column, totalColumn;

  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;
  var date = now.getDate();
  var hour = now.getHours();
  var minute = now.getMinutes();
  var timeOutput = year + '/' + month + '/' + date + ' ' + hour + ':' + minute;

  // ===== Delete old total first =====
  RemoveSummaryRow(sheet, startRow, rangeType === 0 ? 7 : 1);

  if (updateRow === undefined) { // Add new
    if (rangeType === 0) { // Expense
      column = 6; // Number~Note
      totalColumn = 11; // Column K is expense amount total

      // Find last data row (skip total row)
      // Start from last row, search backward
      var lastDataRow = sheet.getLastRow();
      if (lastDataRow < startRow) {
        lastDataRow = startRow - 1; // No data, start from startRow - 1
      } else {
        // Search backward, skip total rows and empty rows
        while (lastDataRow >= startRow) {
          var cellValue = sheet.getRange(lastDataRow, 7).getValue(); // Column G number
          if (cellValue === '總計') {
            lastDataRow--;
            continue;
          }
          if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
            var numValue = Number(cellValue);
            if (!isNaN(numValue) && numValue > 0) {
              break; // Found last data row
            }
          }
          lastDataRow--;
        }
        
        if (lastDataRow < startRow) {
          lastDataRow = startRow - 1; // No valid data found, start from startRow - 1
        }
      }

      // Get last number
      var prev = 0;
      if (lastDataRow >= startRow) {
        prev = sheet.getRange(lastDataRow, 7).getValue(); // Column G number
      }
      var lastNumber = Number(prev);
      if (isNaN(lastNumber) || lastNumber <= 0) {
        lastNumber = 0;
      }

      var row = lastDataRow + 1;
      // Ensure cost is numeric type
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      // Ensure number is numeric type
      var newNumber = lastNumber + 1;
      if (isNaN(newNumber) || newNumber <= 0) {
        newNumber = 1;
      }
      
      // Prepare data to write, ensure all values are not undefined
      var values = [
        newNumber,
        timeOutput || '',
        category || '',
        item || '',
        costValue,
        note || ''
      ];
      
      // Ensure number is in first position (column G) when writing, write 6 columns
      sheet.getRange(row, 7, 1, column).setValues([values]);

      // Add total
      sheet.getRange(row + 1, 7).setValue('總計');
      sheet.getRange(row + 1, 11).setFormula(`=SUM(K${startRow}:K${row})`);

    } else { // Income
      column = 5;
      totalColumn = 4; // Column D is income amount total

      // Find last data row (skip total row)
      // Start from last row, search backward
      var lastDataRow = sheet.getLastRow();
      if (lastDataRow < startRow) {
        lastDataRow = startRow - 1; // No data, start from startRow - 1
      } else {
        // Search backward, skip total rows and empty rows
        while (lastDataRow >= startRow) {
          var cellValue = sheet.getRange(lastDataRow, 1).getValue(); // Column A number
          if (cellValue === '總計') {
            lastDataRow--;
            continue;
          }
          if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
            var numValue = Number(cellValue);
            if (!isNaN(numValue) && numValue > 0) {
              break; // Found last data row
            }
          }
          lastDataRow--;
        }
        
        if (lastDataRow < startRow) {
          lastDataRow = startRow - 1; // No valid data found, start from startRow - 1
        }
      }

      // Get last number
      var prev = 0;
      if (lastDataRow >= startRow) {
        prev = sheet.getRange(lastDataRow, 1).getValue(); // Column A number
      }
      var lastNumber = Number(prev);
      if (isNaN(lastNumber) || lastNumber <= 0) {
        lastNumber = 0;
      }

      var row = lastDataRow + 1;
      // Ensure cost is numeric type
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      // Ensure number is numeric type
      var newNumber = lastNumber + 1;
      if (isNaN(newNumber) || newNumber <= 0) {
        newNumber = 1;
      }
      
      // Prepare data to write, ensure all values are not undefined
      var values = [
        newNumber,
        timeOutput || '',
        item || '',
        costValue,
        note || ''
      ];
      
      // Ensure number is in first position (column A) when writing, write 5 columns
      sheet.getRange(row, 1, 1, column).setValues([values]);

      // Add total
      sheet.getRange(row + 1, 1).setValue('總計');
      sheet.getRange(row + 1, totalColumn).setFormula(`=SUM(D${startRow}:D${row})`);
    }

  } else { // Update data
    if (rangeType === 0) { // Expense
      column = 6;
      totalColumn = 11; // Column K is expense amount total
      
      // Read old value first (for debugging)
      var oldCostValue = sheet.getRange(updateRow, 11).getValue(); // Column K is amount
      var oldCostNum = parseFloat(oldCostValue) || 0;
      
      // ===== Key: clear old value, then update data, finally recalculate total =====
      // Note: RemoveSummaryRow was already called at the start of UpsertData
      
      // 1. First clear the amount field of the old data row (delete entire amount to avoid total formula including old value)
      sheet.getRange(updateRow, 11).clearContent(); // Clear column K amount (important: delete old value first)
      
      // 3. Ensure cost is numeric type
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      
      // 4. Update data (after deleting total row and clearing old value, avoid row number changes and duplicate calculations)
      // updateRow = recordNum + 2 (Row1: type title, Row2: column title), so recordNum = updateRow - 2
      var values = [updateRow - 2, timeOutput, category, item, costValue, note];
      sheet.getRange(updateRow, 7, 1, column).setValues([values]);
      
      // Find last expense data row (skip total row)
      var lastDataRow = sheet.getLastRow();
      while (lastDataRow >= startRow) {
        var cellValue = sheet.getRange(lastDataRow, 7).getValue(); // Column G
        if (cellValue !== '總計' && cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          // Check if it's a numeric number (valid data row)
          var numValue = Number(cellValue);
          if (!isNaN(numValue) && numValue > 0) {
            break; // Found last data row
          }
        }
        lastDataRow--;
      }
      
      if (lastDataRow < startRow) {
        // If no data, set total to 0
        sheet.getRange(startRow, 7).setValue('總計');
        sheet.getRange(startRow, totalColumn).setValue(0);
      } else {
        sheet.getRange(lastDataRow + 1, 7).setValue('總計');
        sheet.getRange(lastDataRow + 1, totalColumn).setFormula(`=SUM(K${startRow}:K${lastDataRow})`);
      }
    } else { // Income
      column = 5;
      totalColumn = 4; // Column D is income amount total
      
      // Read old value first (for debugging)
      var oldCostValue = sheet.getRange(updateRow, 4).getValue(); // Column D is amount
      var oldCostNum = parseFloat(oldCostValue) || 0;
      
      // ===== Key: clear old value, then update data, finally recalculate total =====
      // Note: RemoveSummaryRow was already called at the start of UpsertData
      
      // 1. First clear the amount field of the old data row (delete entire amount to avoid total formula including old value)
      sheet.getRange(updateRow, 4).clearContent(); // Clear column D amount (important: delete old value first)
      
      // 3. Ensure cost is numeric type
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      
      // 4. Update data (after deleting total row and clearing old value, avoid row number changes and duplicate calculations)
      // updateRow = recordNum + 2 (Row1: type title, Row2: column title), so recordNum = updateRow - 2
      var values = [updateRow - 2, timeOutput, item, costValue, note];
      sheet.getRange(updateRow, 1, 1, column).setValues([values]);
      
      // Find last income data row (skip total row)
      var lastDataRow = sheet.getLastRow();
      while (lastDataRow >= startRow) {
        var cellValue = sheet.getRange(lastDataRow, 1).getValue(); // Column A
        if (cellValue !== '總計' && cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          // Check if it's a numeric number (valid data row)
          var numValue = Number(cellValue);
          if (!isNaN(numValue) && numValue > 0) {
            break; // Found last data row
          }
        }
        lastDataRow--;
      }
      
      if (lastDataRow < startRow) {
        // If no data, set total to 0
        sheet.getRange(startRow, 1).setValue('總計');
        sheet.getRange(startRow, totalColumn).setValue(0);
      } else {
        sheet.getRange(lastDataRow + 1, 1).setValue('總計');
        sheet.getRange(lastDataRow + 1, totalColumn).setFormula(`=SUM(D${startRow}:D${lastDataRow})`);
      }
    }
  }

  // Invalidate cache after data modification
  invalidateCache(sheetIndex);

  return { success: true, message: '資料已成功新增', data: ShowTabData(sheetIndex), total: GetSummary(sheetIndex) };
}

// ===== Supplement: clear old total row (only clear specific block columns, don't delete entire row) =====
function RemoveSummaryRow(sheet, startRow, labelCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  
  // Find the summary row for this block (search from bottom up within the block)
  for (var r = lastRow; r >= startRow; r--) {
    var cellValue = sheet.getRange(r, labelCol).getValue();
    if (cellValue === '總計') {
      // Clear only the specific columns for this block, not the entire row
      if (labelCol === 1) {
        // Income block: columns A-E (1-5)
        sheet.getRange(r, 1, 1, 5).clearContent();
      } else if (labelCol === 7) {
        // Expense block: columns G-L (7-12)
        sheet.getRange(r, 7, 1, 6).clearContent();
      }
      return;
    }
  }
}


function DeleteTab(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var sheetName = targetSheet.getSheetName();
  ss.deleteSheet(targetSheet);

  return { success: true, message: '分頁已成功刪除: ' + sheetName };
}

function DeleteData(sheetIndex, rangeType, number) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 3; // 從第三行開始（Row1: 類型標題, Row2: 欄位標題, Row3: 資料開始）
  var column, totalColumn, startCol, numberCol;

  if (rangeType === 0) { // Expense
    column = 6;       // Number~Note (6 columns)
    totalColumn = 11; // Column K is expense amount total
    numberCol = 7;    // Column G is number
    startCol = 7;     // Expense starts at column G
  } else { // Income
    column = 5;       // Number~Note (5 columns)
    totalColumn = 4;  // Column D is income amount total
    numberCol = 1;    // Column A is number
    startCol = 1;     // Income starts at column A
  }

  // Find all data rows for this block (to find the one to delete and shift others up)
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  var targetNumber = parseInt(number, 10);
  var dataRows = []; // Store all data rows for this block
  var summaryRow = -1;
  
  for (var r = startRow; r <= lastRow; r++) {
    var cellVal = sheet.getRange(r, numberCol).getValue();
    if (cellVal === '總計') {
      summaryRow = r;
      continue;
    }
    var cellNum = parseInt(cellVal, 10);
    if (!isNaN(cellNum) && cellNum > 0) {
      dataRows.push({ row: r, num: cellNum });
      if (cellNum === targetNumber) {
        targetRow = r;
      }
    }
  }

  if (targetRow === -1) {
    return { success: false, message: 'Cannot find corresponding number data' };
  }

  // Find the index of the target row in dataRows
  var targetIndex = -1;
  for (var i = 0; i < dataRows.length; i++) {
    if (dataRows[i].row === targetRow) {
      targetIndex = i;
      break;
    }
  }

  // Shift data up: copy each row below the target to one row above
  for (var i = targetIndex; i < dataRows.length - 1; i++) {
    var currentRow = dataRows[i].row;
    var nextRow = dataRows[i + 1].row;
    var nextData = sheet.getRange(nextRow, startCol, 1, column).getValues()[0];
    sheet.getRange(currentRow, startCol, 1, column).setValues([nextData]);
  }

  // Clear the last data row (it's now a duplicate)
  if (dataRows.length > 0) {
    var lastDataRowNum = dataRows[dataRows.length - 1].row;
    sheet.getRange(lastDataRowNum, startCol, 1, column).clearContent();
  }

  // Clear old summary row if exists
  if (summaryRow > 0) {
    sheet.getRange(summaryRow, startCol, 1, column).clearContent();
  }

  // Renumber the remaining records and find the new last data row
  var newLastDataRow = -1;
  var newNumber = 1;
  for (var i = 0; i < dataRows.length - 1; i++) { // -1 because we deleted one
    var rowNum = dataRows[i].row;
    sheet.getRange(rowNum, numberCol).setValue(newNumber);
    newLastDataRow = rowNum;
    newNumber++;
  }

  // Set new total row
  if (newLastDataRow < startRow) {
    // No data left, set total to 0 at startRow
    sheet.getRange(startRow, numberCol).setValue('總計');
    sheet.getRange(startRow, totalColumn).setValue(0);
  } else {
    // Add total row after the last data row
    var newSummaryRow = newLastDataRow + 1;
    sheet.getRange(newSummaryRow, numberCol).setValue('總計');
    sheet.getRange(newSummaryRow, totalColumn).setFormula(
      `=SUM(${sheet.getRange(startRow, totalColumn).getA1Notation()}:${sheet.getRange(newLastDataRow, totalColumn).getA1Notation()})`
    );
  }

  // Invalidate cache after data modification
  invalidateCache(sheetIndex);

  return {
    success: true,
    message: 'Data successfully deleted, total updated',
    data: ShowTabData(sheetIndex),
    total: GetSummary(sheetIndex)
  };
}

function ChangeTabName(sheet,name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var oldName = targetSheet.getSheetName();
  targetSheet.setName(name);

  return { success: true, message: 'Tab name changed from "' + oldName + '" to "' + name + '"' };
}

function GetSummary(sheet){
  var cacheKey = getCacheKey('summary', sheet);
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var income = 0;
  var expense = 0;
  var incomeFound = false;
  var expenseFound = false;

  // Only read rows that have data (not entire column)
  var lastRow = targetSheet.getLastRow();
  if (lastRow > 0) {
    // Read both income (A:D) and expense (G:K) in one call
    var data = targetSheet.getRange(1, 1, lastRow, 11).getValues();

    // Search from back to front, find the last "總計" row (should be the latest)
    for (var i = data.length - 1; i >= 0; i--) {
      // Find income total: look for "總計" in column A, get value from column D
      if (data[i][0] === '總計' && !incomeFound) {
        // Read value directly from column D (index 3)
        var incomeCell = targetSheet.getRange(i + 1, 4); // Column D is column 4
        var incomeValue = incomeCell.getValue();
        var incomeFormula = incomeCell.getFormula();
        var incomeDisplayValue = incomeCell.getDisplayValue();
        
        // Debug: log the raw value read
        Logger.log('[GetSummary] 收入總計行: ' + (i + 1) + ', D欄值: ' + incomeValue + ', 公式: ' + incomeFormula + ', 顯示值: ' + incomeDisplayValue);
        
        // If there's a formula, force recalculation and read display value
        if (incomeFormula && incomeFormula.trim() !== '') {
          if (incomeDisplayValue && incomeDisplayValue.trim() !== '') {
            var cleanValue = incomeDisplayValue.replace(/,/g, '').trim();
            incomeValue = parseFloat(cleanValue);
            if (isNaN(incomeValue)) {
              incomeValue = incomeCell.getValue(); // Fall back to getValue()
            }
            Logger.log('[GetSummary] 從顯示值解析收入: ' + incomeValue);
          } else {
            incomeValue = incomeCell.getValue();
            Logger.log('[GetSummary] 顯示值為空，使用 getValue(): ' + incomeValue);
          }
        }
        
        // Ensure value is numeric type
        if (typeof incomeValue === 'number' && !isNaN(incomeValue)) {
          income = incomeValue;
          Logger.log('[GetSummary] 收入（數字類型）: ' + income);
        } else if (typeof incomeValue === 'string' && incomeValue.trim() !== '') {
          var cleanValue = incomeValue.replace(/,/g, '').trim();
          income = parseFloat(cleanValue) || 0;
          Logger.log('[GetSummary] 收入（字串解析）: ' + income);
        } else {
          income = 0;
          Logger.log('[GetSummary] 收入設為 0（無法解析）');
        }
        
        // No matter what value is read, try to manually calculate income (ensure accuracy)
        // Check if there are data rows before the total row
        if (i + 1 > 2) {
          var hasIncomeData = false;
          for (var checkRow = 2; checkRow < i + 1; checkRow++) {
            var checkNumber = targetSheet.getRange(checkRow, 1).getValue(); // Column A number
            if (typeof checkNumber === 'number' && checkNumber > 0) {
              hasIncomeData = true;
              if (typeof Logger !== 'undefined') {
                Logger.log('[GetSummary] 找到收入數據行: ' + checkRow);
              }
              break;
            }
          }
          
          if (hasIncomeData) {
            var incomeStartRow = 2;
            var incomeSum = 0;
            var incomeRowCount = 0;
            for (var r = incomeStartRow; r < i + 1; r++) {
              var incomeRowData = targetSheet.getRange(r, 1, 1, 5).getValues()[0]; // Columns A to E
              var rowNumber = incomeRowData[0];
              if (typeof rowNumber === 'number' && rowNumber > 0) {
                var rowCost = incomeRowData[3]; // Column D is amount (index 3)
                if (typeof Logger !== 'undefined') {
                  Logger.log('[GetSummary] 收入行 ' + r + ': 編號=' + rowNumber + ', 金額=' + rowCost + ' (類型: ' + typeof rowCost + ')');
                }
                if (typeof rowCost === 'number' && !isNaN(rowCost)) {
                  incomeSum += rowCost;
                  incomeRowCount++;
                } else if (typeof rowCost === 'string' && rowCost.trim() !== '') {
                  var numCost = parseFloat(rowCost.replace(/,/g, '')) || 0;
                  incomeSum += numCost;
                  incomeRowCount++;
                }
              }
            }
            if (typeof Logger !== 'undefined') {
              Logger.log('[GetSummary] 手動計算收入總計: ' + incomeSum + ' (共 ' + incomeRowCount + ' 筆記錄)');
              Logger.log('[GetSummary] 讀取到的收入值: ' + income);
            }
            // If manually calculated result is not 0, prioritize using manually calculated result
            // If manually calculated result is 0 but read value is not 0, use read value
            if (incomeSum !== 0) {
              income = incomeSum;
              if (typeof Logger !== 'undefined') {
                Logger.log('[GetSummary] 使用手動計算的收入: ' + income);
              }
            } else if (income === 0 && incomeSum === 0 && incomeRowCount > 0) {
              // If there are data rows but calculation result is 0, all amounts might be 0, which is normal
              if (typeof Logger !== 'undefined') {
                Logger.log('[GetSummary] 收入數據行存在但金額總和為 0');
              }
            }
          } else {
            if (typeof Logger !== 'undefined') {
              Logger.log('[GetSummary] 沒有找到收入數據行');
            }
          }
        }
        
        incomeFound = true;
      }
      // Find expense total: look for "總計" in column G, get value from column K
      if (data[i][6] === '總計' && !expenseFound) {
        // Read value directly from column K (index 10, corresponds to column 11)
        var expenseCell = targetSheet.getRange(i + 1, 11); // Column K is column 11
        var expenseValue = expenseCell.getValue();
        var cellFormula = expenseCell.getFormula();
        
        // If there's a formula, force recalculation and read display value
        if (cellFormula && cellFormula.trim() !== '') {
          // Use getDisplayValue() to get the displayed value after formula calculation
          var displayValue = expenseCell.getDisplayValue();
          if (displayValue && displayValue.trim() !== '') {
            // Remove thousand separators and other format characters
            var cleanValue = displayValue.replace(/,/g, '').trim();
            expenseValue = parseFloat(cleanValue);
            if (isNaN(expenseValue)) {
              expenseValue = expenseCell.getValue(); // Fall back to getValue()
            }
          } else {
            // If display value is empty, try using getValue()
            expenseValue = expenseCell.getValue();
          }
        }
        
        // Ensure value is numeric type
        if (typeof expenseValue === 'number' && !isNaN(expenseValue)) {
          expense = expenseValue;
        } else if (typeof expenseValue === 'string' && expenseValue.trim() !== '') {
          // Remove thousand separators and other format characters
          var cleanValue = expenseValue.replace(/,/g, '').trim();
          expense = parseFloat(cleanValue) || 0;
        } else {
          expense = 0;
        }
        
        // If read value is 0 but formula exists, and there are data rows before total row, try manual calculation
        if (expense === 0 && cellFormula && cellFormula.trim() !== '' && i + 1 > 2) {
          // Check if there are expense records (from row 2 to before total row)
          var hasExpenseData = false;
          for (var checkRow = 2; checkRow < i + 1; checkRow++) {
            var checkNumber = targetSheet.getRange(checkRow, 7).getValue(); // Column G number
            if (typeof checkNumber === 'number' && checkNumber > 0) {
              hasExpenseData = true;
              break;
            }
          }
          
          // If there are expense records but total is 0, manually calculate
          if (hasExpenseData) {
            var expenseStartRow = 2; // Data starts from row 2
            var expenseSum = 0;
            for (var r = expenseStartRow; r < i + 1; r++) {
              var expenseRowData = targetSheet.getRange(r, 7, 1, 6).getValues()[0]; // Columns G to L
              // Check if it's a valid expense record (column G should be a numeric number)
              var rowNumber = expenseRowData[0];
              if (typeof rowNumber === 'number' && rowNumber > 0) {
                // Column K is amount (index 4)
                var rowCost = expenseRowData[4];
                if (typeof rowCost === 'number' && !isNaN(rowCost)) {
                  expenseSum += rowCost;
                } else if (typeof rowCost === 'string' && rowCost.trim() !== '') {
                  var numCost = parseFloat(rowCost.replace(/,/g, '')) || 0;
                  expenseSum += numCost;
                }
              }
            }
            if (expenseSum !== 0) {
              expense = expenseSum;
            }
          }
        }
        
        expenseFound = true;
      }
      // Stop if both found
      if (incomeFound && expenseFound) break;
    }
  }

  var total = income - expense;
  var result = [income, expense, total];
  
  // Debug: log final result (use Logger if available, otherwise use console)
  if (typeof Logger !== 'undefined') {
    Logger.log('[GetSummary] 最終結果: income=' + income + ', expense=' + expense + ', total=' + total);
    Logger.log('[GetSummary] 結果陣列: [' + result.join(', ') + ']');
  }
  
  setToCache(cacheKey, result);
  return result;
}

function arrayMatch(a1, a2) {
  if (!Array.isArray(a1) || !Array.isArray(a2)) return false;
  if (a1.length !== a2.length) return false;
  for (let i = 0; i < a1.length; i++) {
    const val1 = a1[i];
    const val2 = a2[i];
    const bothArrays = Array.isArray(val1) && Array.isArray(val2);
    if (bothArrays) {
      if (!arrayMatch(val1, val2)) return false;
    } else if (val1 !== val2) {
      return false;
    }
  }
  return true;
}