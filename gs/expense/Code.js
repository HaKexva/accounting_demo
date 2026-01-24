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
  // New: receive date field (optional)
  var dateValue = contents.date;   // e.g. "2025/01/31" or "2025-01-31"
  var item = contents.item;
  var category = contents.category;
  var spendWay = contents.spendWay;
  var creditCard = contents.creditCard
  var month = contents.month
  var actualCost = contents.actualCost
  var payment = contents.payment
  var recordCost = contents.recordCost
  var note = contents.note;
  var updateRow = contents.updateRow;
  var oldValue = contents.oldValue
  var newValue = contents.newValue
  var itemId = contents.itemId
  var action = contents.action
  // New: receive drag-and-drop sorting indices
  var oldIndex = contents.oldIndex;
  var newIndex = contents.newIndex;
  // New: for batch updates
  var originalData = contents.originalData;
  var newData = contents.newData;

  var result = { success: false, message: "" };

  try {
    if (name === "Upsert Data") {
      // Pass dateValue so UpsertData can use the specified date
      result = UpsertData(sheet, dateValue, item, category, spendWay, creditCard, month, actualCost, payment, recordCost, note, updateRow);
    } else if (name === "Create Tab") {
      result = CreateNewTab();
    } else if (name === "Delete Data") {
      result = DeleteData(sheet,updateRow);
    } else if (name === "Update Dropdown"){
      result = UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex);
    } else if (name === "Batch Update Dropdown") {
      result = BatchUpdateDropdown(itemId, originalData, newData);
    } else {
      result.message = "Unknown operation type: " + name;
    }
  } catch (error) {
    result.success = false;
    result.message = "Operation failed: " + error.toString();
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
  allSheets.slice(2).forEach(function(sheet){
    sheetNames.push(sheet.getSheetName());
  });

  setToCache(cacheKey, sheetNames);
  return sheetNames;
}

function ShowTabData(sheet) {
  var cacheKey = getCacheKey('tabData', sheet);
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var lastRow = targetSheet.getRange('A' + targetSheet.getMaxRows() + ':J' + targetSheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow()
  var output = targetSheet.getRange('A1:J' + lastRow).getValues()

  setToCache(cacheKey, output);
  return output
}

function CreateNewTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var targetSheet = ss.getSheets()[0];
  var value = [["Total",,,,,,"0",,"0"]]
  targetSheet.getRange("A2:I2").setValues(value)
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
    console.log('Next month sheet already created, please do not add again');
    return { success: false, message: 'Next month sheet already created, please do not add again' };
  } else {
    var destination = SpreadsheetApp.openById('1pOdDjlyhyyWpbpgumrorgUgT2VTauXdq8bt92Et0CoA');
    targetSheet.copyTo(destination);
    var copyName = ss.getSheetByName('「空白表」的副本');
    copyName.setName(year + month.toString());
  }
  // 清除月份列表快取，確保下次載入時取得最新列表
  invalidateTabNamesCache();
  return { success: true, message: 'New tab successfully created: ' + year + month };
}

function CleanupSummary(sheetIndex) {  // Changed to receive sheetIndex
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];  // Use sheetIndex to get sheet
  var column = 10
  var row = sheet.getRange('A' + sheet.getMaxRows() + ':J' + sheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (sheet.getRange(row,1).getValue() === 'Total') {
    sheet.getRange(row, 1, 1, column).clear();
  }
}

function UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dropdownSheet = ss.getSheets()[1];
  var data = dropdownSheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, message: 'Dropdown table is empty' };
  }

  var headerRow = data[0];
  var colIndex = -1;

  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['消費類別', '類別']);
  } else if (itemId === 'payment') {
    colIndex = FindHeaderColumn(headerRow, ['支付方式']);
  } else if (itemId === 'platform') {
    colIndex = FindHeaderColumn(headerRow, ['支付平台', '平台']);
  }

  if (colIndex < 0) {
    return { success: false, message: 'Cannot find corresponding column' };
  }

  // === New: Drag-and-drop sorting feature ===
  if (action === 'reorder') {
    // Check index parameters
    if (typeof oldIndex !== 'number' || typeof newIndex !== 'number') {
      return { success: false, message: 'When sorting, oldIndex and newIndex must be numbers' };
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
      return { success: false, message: 'Index out of range' };
    }

    // Execute sorting: move item from oldIndex to newIndex position
    var movedItem = values.splice(oldIndex, 1)[0];
    values.splice(newIndex, 0, movedItem);

    // Update dropdown table (write starting from row 2) - batch write all values at once
    if (values.length > 0) {
      var writeData = values.map(function(v) { return [v]; });
      dropdownSheet.getRange(2, colIndex + 1, values.length, 1).setValues(writeData);
    }

    // Clear remaining extra cells
    var writeRow = 2 + values.length;
    if (writeRow <= data.length) {
      dropdownSheet.getRange(writeRow, colIndex + 1, data.length - writeRow + 1, 1).clearContent();
    }

    return { success: true, message: 'Sorting successful, dropdown order updated' };
  }

  var updatedCount = 0;

  // === New: Find the last row in this column and add to the next row ===
  if (action === 'add') {
    if (!newValue) {
      return { success: false, message: 'newValue is required when adding' };
    }

    // Find the last row with value in this column (search from bottom to top)
    var lastRowInColumn = 1; // Default to header row (index 1, corresponds to row 2)
    for (var r = data.length - 1; r >= 1; r--) { // Search from last row upward, skip header row
      var val = data[r][colIndex];
      if (val !== '' && val !== null && val !== undefined) {
        lastRowInColumn = r + 1; // Found the last row with value (+1 to convert to row number)
        break;
      }
    }

    // Add at the next position after the last row in this column
    dropdownSheet.getRange(lastRowInColumn + 1, colIndex + 1).setValue(newValue);
    return { success: true, message: 'Added to dropdown', updated: 1 };
  }

  // === Delete: Only update dropdown table ===
  if (action === 'delete') {
    if (!oldValue) {
      return { success: false, message: 'oldValue is required when deleting (cannot be empty)' };
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
    return { success: true, message: 'Deleted from dropdown', updated: updatedCount };
  }

  // === Edit (merge): Update dropdown table + historical data in all monthly sheets ===
  if (action === 'edit') {
    if (!oldValue || !newValue || oldValue === newValue) {
      return { success: false, message: 'oldValue/newValue incorrect when editing (oldValue cannot be empty)' };
    }

    // 1. Update dropdown table
    for (var r2 = 1; r2 < data.length; r2++) {
      var val2 = data[r2][colIndex];
      if (val2 === oldValue) {
        data[r2][colIndex] = newValue;
        updatedCount++;
      }
    }
    dropdownSheet.getDataRange().setValues(data);

    // 2. Update historical data in all monthly sheets (starting from index 2)
    var allSheets = ss.getSheets();
    var monthSheetsUpdated = 0;
    var totalRecordsUpdated = 0;

    // Determine the column index to update (based on UpsertData's column order)
    // Expense table column order: [Date, Item, Category, Payment Method, Credit Card Payment Method, This/Next Month Payment, Actual Cost, Payment Platform, Recorded Cost, Notes]
    var targetColumnIndex = -1;
    if (itemId === 'category') {
      targetColumnIndex = 2; // Column 3 (index 2): Category
    } else if (itemId === 'payment') {
      targetColumnIndex = 3; // Column 4 (index 3): Payment Method
    } else if (itemId === 'platform') {
      targetColumnIndex = 7; // Column 8 (index 7): Payment Platform
    }

    if (targetColumnIndex >= 0) {
      for (var s = 2; s < allSheets.length; s++) { // Start from index 2 (skip first two sheets)
        var monthSheet = allSheets[s];
        var monthData = monthSheet.getDataRange().getValues();

        if (monthData.length < 2) continue; // Skip empty sheets

        var monthUpdated = false;

        // Start from row 2 (index 1), skip header row
        for (var row = 1; row < monthData.length; row++) {
          // Check if it's the "Total" row or empty row
          if (monthData[row][0] === 'Total' || monthData[row][0] === '總計' || monthData[row][0] === '') {
            continue;
          }

          var cellValue = monthData[row][targetColumnIndex];
          if (cellValue === oldValue) {
            monthSheet.getRange(row + 1, targetColumnIndex + 1).setValue(newValue);
            monthUpdated = true;
            totalRecordsUpdated++;
          }
        }

        if (monthUpdated) {
          monthSheetsUpdated++;
        }
      }
    }

    return {
      success: true,
      message: 'Merge completed. Dropdown updated: ' + updatedCount + ' items. Monthly sheets updated: ' + totalRecordsUpdated + ' records (' + monthSheetsUpdated + ' months)'
    };
  }

  return { success: false, message: 'Unknown action: ' + action };
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

// Adjusted parameter order: added dateValue
function UpsertData(sheetIndex, dateValue, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note, updateRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 2;
  var values;

  // Default to "today's" date string
  var now = new Date();
  var year = now.getFullYear().toString();
  var month = (now.getMonth() + 1).toString();
  var day = now.getDate().toString();
  var defaultDateString = year + '/' + month + '/' + day;

  // If frontend passed date, use it; otherwise use defaultDateString
  var timeOutput = dateValue ? dateValue : defaultDateString;

  if (updateRow === undefined) { // Add new
    CleanupSummary(sheetIndex);
    var lastDataRow = getLastDataRow(sheet, startRow);
    var insertRow = lastDataRow + 1;

    // First column is date
    values = [timeOutput, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note];
    sheet.getRange(insertRow, 1, 1, 10).setValues([values]);

    var totalRow = insertRow + 1;
    sheet.getRange(totalRow, 1).setValue('Total');

    var firstCell = sheet.getRange(startRow, 7).getA1Notation();
    var lastCell = sheet.getRange(insertRow, 7).getA1Notation();
    sheet.getRange(totalRow, 7).setFormula(`=SUM(${firstCell}:${lastCell})`);

    var firstCell2 = sheet.getRange(startRow, 9).getA1Notation();
    var lastCell2 = sheet.getRange(insertRow, 9).getA1Notation();
    sheet.getRange(totalRow, 9).setFormula(`=SUM(${firstCell2}:${lastCell2})`);

  } else { // Update
    // If frontend passed new dateValue, update the date; otherwise keep the original date.
    var existingDate = sheet.getRange(updateRow, 1).getValue();
    var finalDate = dateValue ? dateValue : existingDate;

    values = [finalDate, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note];
    sheet.getRange(updateRow, 1, 1, 10).setValues([values]);
  }

  // Invalidate cache after data modification
  invalidateCache(sheetIndex);

  return { success: true, message: 'Data successfully saved', data: ShowTabData(sheetIndex), total: GetSummary(sheetIndex) };
}

function DeleteData(sheetIndex, updateRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 2;

  // Avoid deleting header or invalid row
  if (!updateRow || updateRow < startRow || updateRow > sheet.getLastRow()) {
    return { success: false, message: 'Invalid row to delete: ' + updateRow };
  }

  CleanupSummary(sheetIndex);
  sheet.deleteRow(updateRow);

  var lastDataRow = getLastDataRow(sheet, startRow);
  var totalRow = lastDataRow + 1;
  sheet.getRange(totalRow, 1).setValue('Total');

  if (lastDataRow >= startRow) {
    var firstCell = sheet.getRange(startRow, 7).getA1Notation();
    var lastCell = sheet.getRange(lastDataRow, 7).getA1Notation();
    sheet.getRange(totalRow, 7).setFormula(`=SUM(${firstCell}:${lastCell})`);

    var firstCell2 = sheet.getRange(startRow, 9).getA1Notation();
    var lastCell2 = sheet.getRange(lastDataRow, 9).getA1Notation();
    sheet.getRange(totalRow, 9).setFormula(`=SUM(${firstCell2}:${lastCell2})`);
  }

  // Invalidate cache after data modification
  invalidateCache(sheetIndex);

  return {
    success: true,
    message: 'Data successfully deleted',
    data: ShowTabData(sheetIndex),
    total: GetSummary(sheetIndex)
  };
}

function getLastDataRow(sheet, startRow) {
  var lastRow = sheet.getLastRow();

  // No data at all (only header)
  if (lastRow < startRow) {
    return startRow - 1;
  }

  // Batch read all values in column A at once (much faster than row-by-row)
  var numRows = lastRow - startRow + 1;
  var values = sheet.getRange(startRow, 1, numRows, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    var val = values[i][0];
    if (val === '' || val === 'Total' || val === '總計' || val === null) {
      return startRow + i - 1;
    }
  }

  return lastRow;
}

function GetSummary(sheet){
  var cacheKey = getCacheKey('summary', sheet);
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var actualCost = 0;
  var recordCost = 0;

  // Only read rows that have data (not entire column)
  var lastRow = targetSheet.getLastRow();
  if (lastRow > 0) {
    var data = targetSheet.getRange(1, 1, lastRow, 10).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === 'Total' || data[i][0] === '總計') {
        actualCost = data[i][6] || 0;  // Column G (index 6)
        recordCost = data[i][8] || 0;  // Column I (index 8)
        break;
      }
    }
  }

  var result = [actualCost, recordCost];
  setToCache(cacheKey, result);
  return result;
}

// Batch update dropdown
// @param {string} itemId - Item ID (category, payment, platform)
// @param {Array} originalData - Original data array
// @param {Array} newData - New data array
function BatchUpdateDropdown(itemId, originalData, newData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dropdownSheet = ss.getSheets()[1]; // Dropdown table
  var data = dropdownSheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, message: 'Dropdown table is empty' };
  }

  // Find the corresponding column
  var headerRow = data[0];
  var colIndex = -1;

  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['消費類別', '類別']);
  } else if (itemId === 'payment') {
    colIndex = FindHeaderColumn(headerRow, ['支付方式']);
  } else if (itemId === 'platform') {
    colIndex = FindHeaderColumn(headerRow, ['支付平台', '平台']);
  }

  if (colIndex < 0) {
    return { success: false, message: 'Cannot find corresponding column: ' + itemId };
  }

  // Validate data
  if (!Array.isArray(originalData) || !Array.isArray(newData)) {
    return { success: false, message: 'originalData and newData must be arrays' };
  }

  // Calculate changes
  var originalSet = new Set(originalData);
  var newSet = new Set(newData);

  var removed = originalData.filter(function(item) { return !newSet.has(item); });
  var added = newData.filter(function(item) { return !originalSet.has(item); });

  // Determine if it's a rename (removed count equals added count)
  var renames = [];
  if (removed.length > 0 && removed.length === added.length) {
    for (var i = 0; i < removed.length; i++) {
      renames.push({ oldValue: removed[i], newValue: added[i] });
    }
  }

  // Clear all values in this column (keep header)
  var maxRows = dropdownSheet.getMaxRows();
  if (maxRows > 1) {
    dropdownSheet.getRange(2, colIndex + 1, maxRows - 1, 1).clearContent();
  }

  // Write new data - batch write all values at once
  if (newData.length > 0) {
    var writeData = newData.map(function(v) { return [v]; });
    dropdownSheet.getRange(2, colIndex + 1, newData.length, 1).setValues(writeData);
  }

  // If it's "category" and there are renames, update historical data in all monthly sheets
  var totalRecordsUpdated = 0;
  var monthSheetsUpdated = 0;

  if (itemId === 'category' && renames.length > 0) {
    var allSheets = ss.getSheets();

    // Expense table column order: [Date, Item, Category, Payment Method, Credit Card Payment Method, This/Next Month Payment, Actual Cost, Payment Platform, Recorded Cost, Notes]
    var targetColumnIndex = 2; // Column 3 (index 2): Category

    for (var s = 2; s < allSheets.length; s++) { // Start from index 2 (skip first two sheets)
      var monthSheet = allSheets[s];
      var monthData = monthSheet.getDataRange().getValues();

      if (monthData.length < 2) continue; // Skip empty sheets

      var monthUpdated = false;

      // Start from row 2 (index 1), skip header row
      for (var row = 1; row < monthData.length; row++) {
        // Check if it's the "Total" row or empty row
        if (monthData[row][0] === 'Total' || monthData[row][0] === '總計' || monthData[row][0] === '') {
          continue;
        }

        var cellValue = monthData[row][targetColumnIndex];

        // Check if rename is needed
        for (var r = 0; r < renames.length; r++) {
          if (cellValue === renames[r].oldValue) {
            monthSheet.getRange(row + 1, targetColumnIndex + 1).setValue(renames[r].newValue);
            monthUpdated = true;
            totalRecordsUpdated++;
            break;
          }
        }
      }

      if (monthUpdated) {
        monthSheetsUpdated++;
      }
    }
  }

  var message = 'Batch update completed.';
  if (removed.length > 0) message += ' Removed: ' + removed.length + ' items.';
  if (added.length > 0) message += ' Added: ' + added.length + ' items.';
  if (renames.length > 0) message += ' Renamed: ' + renames.length + ' items.';
  if (totalRecordsUpdated > 0) message += ' Historical records updated: ' + totalRecordsUpdated + ' records (' + monthSheetsUpdated + ' months).';

  return {
    success: true,
    message: message,
    removed: removed.length,
    added: added.length,
    renamed: renames.length,
    recordsUpdated: totalRecordsUpdated
  };
}
