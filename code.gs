/**
 * 📊 StockFlow - Enterprise Central Store & Material Allotment Register
 * Google Apps Script Web App Code (Code.gs)
 * 
 * Provides robust Google Sheets API endpoints for complete data persistence
 * of Inventory, Allotments, Procurements, Scrap, and System Logs.
 */

// 1. WEB APP SERVICE HANDLER
function doGet(e) {
  // Serves the index.html with iframe framing protection lifted for embedding
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('StockFlow - ITI Material Allotment & Tracking System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include HTML components/css if separated (used in multi-file GAS layouts)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 2. DATABASE INITIALIZATION & SHEET MANAGEMENT
function getOrCreateDatabase() {
  var dbs = DriveApp.getFilesByName("StockFlow_Database");
  var ss;
  if (dbs.hasNext()) {
    ss = SpreadsheetApp.open(dbs.next());
  } else {
    ss = SpreadsheetApp.create("StockFlow_Database");
    
    // Create initial schemas
    var inventorySheet = ss.insertSheet("Inventory");
    inventorySheet.appendRow(["ID", "Name", "Category", "Subcategory", "SKU", "Quantity", "Unit", "LowStockThreshold", "AvgUnitPrice", "ShelfLocation", "Brand", "GeMCategory", "PurchaseDate", "ExpiryDate", "CreatedDate"]);
    
    var allotmentsSheet = ss.insertSheet("Allotments");
    allotmentsSheet.appendRow(["ID", "RequesterName", "ItemsJSON", "Purpose", "Status", "Level1Status", "Level2Status", "Level3Status", "CreatedDate"]);
    
    var procurementSheet = ss.insertSheet("Procurements");
    procurementSheet.appendRow(["ID", "ItemName", "Quantity", "Unit", "Category", "EstimatedCost", "Purpose", "Status", "Urgency", "RequestedBy", "ApprovedBy", "Comments", "CreatedDate"]);
    
    var scrapSheet = ss.insertSheet("Scrap");
    scrapSheet.appendRow(["ID", "ItemName", "Quantity", "Unit", "Category", "DisposedDate", "Valuation", "Method", "ApprovedBy", "Status", "Reason", "Notes"]);
    
    var logsSheet = ss.insertSheet("Logs");
    logsSheet.appendRow(["ID", "Timestamp", "User", "Role", "Action", "Details", "Type"]);
    
    // Delete default "Sheet1"
    var defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  }
  return ss;
}

// 3. COLLECTION READ API
function readCollection(sheetName) {
  try {
    var ss = getOrCreateDatabase();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return JSON.stringify([]);
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify([]);
    
    var headers = data[0];
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        // CamelCase the headers: "LowStockThreshold" -> "lowStockThreshold"
        var formattedKey = key.charAt(0).toLowerCase() + key.slice(1);
        var val = row[j];
        
        // Parse JSON fields automatically
        if (formattedKey === 'itemsJSON' || formattedKey === 'items') {
          try {
            obj['items'] = JSON.parse(val);
          } catch(e) {
            obj['items'] = [];
          }
        } else {
          obj[formattedKey] = val;
        }
      }
      results.push(obj);
    }
    return JSON.stringify(results);
  } catch (error) {
    return JSON.stringify({ error: error.message });
  }
}

// 4. COLLECTION BULK SAVE/SYNC API
function saveCollection(sheetName, jsonArray) {
  try {
    var list = JSON.parse(jsonArray);
    var ss = getOrCreateDatabase();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return JSON.stringify({ error: "Sheet not found" });
    
    // Clear and rebuild headers
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    sheet.clearContents();
    sheet.appendRow(headers);
    
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      var row = [];
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        var formattedKey = key.charAt(0).toLowerCase() + key.slice(1);
        
        var val = item[formattedKey];
        if (formattedKey === 'itemsJSON' || formattedKey === 'items') {
          val = JSON.stringify(item['items'] || item['itemsJSON'] || []);
        } else if (val === undefined || val === null) {
          val = "";
        }
        row.push(val);
      }
      sheet.appendRow(row);
    }
    return JSON.stringify({ success: true, count: list.length });
  } catch (error) {
    return JSON.stringify({ error: error.message });
  }
}

// 5. INDIVIDUAL WRITE/UPDATE HANDLERS
function updateItemInCollection(sheetName, itemJson) {
  try {
    var item = JSON.parse(itemJson);
    var ss = getOrCreateDatabase();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return JSON.stringify({ error: "Sheet not found" });
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var foundRowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == item.id) {
        foundRowIndex = i + 1; // 1-indexed for Sheets
        break;
      }
    }
    
    var rowValues = [];
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var formattedKey = key.charAt(0).toLowerCase() + key.slice(1);
      
      var val = item[formattedKey];
      if (formattedKey === 'itemsJSON' || formattedKey === 'items') {
        val = JSON.stringify(item['items'] || item['itemsJSON'] || []);
      } else if (val === undefined || val === null) {
        val = "";
      }
      rowValues.push(val);
    }
    
    if (foundRowIndex > -1) {
      sheet.getRange(foundRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
    return JSON.stringify({ success: true, id: item.id });
  } catch (error) {
    return JSON.stringify({ error: error.message });
  }
}

// 6. WRAPPERS TO MATCH STOCKFLOW API EXPECTATIONS
function getInventory() { return readCollection("Inventory"); }
function getAllotments() { return readCollection("Allotments"); }
function getProcurements() { return readCollection("Procurements"); }
function getScrap() { return readCollection("Scrap"); }
function getLogs() { return readCollection("Logs"); }

function saveInventoryList(json) { return saveCollection("Inventory", json); }
function saveAllotmentsList(json) { return saveCollection("Allotments", json); }
function saveProcurementsList(json) { return saveCollection("Procurements", json); }
function saveScrapList(json) { return saveCollection("Scrap", json); }
function saveLogsList(json) { return saveCollection("Logs", json); }

function updateInventoryItem(json) { return updateItemInCollection("Inventory", json); }
function saveAllotmentRequest(json) { return updateItemInCollection("Allotments", json); }
function updateProcurementRequest(json) { return updateItemInCollection("Procurements", json); }
function updateScrapRecord(json) { return updateItemInCollection("Scrap", json); }
function addSystemLog(json) { return updateItemInCollection("Logs", json); }
