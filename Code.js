function doGet(e) {
  if (!e.parameter.action) {
    return ContentService.createTextOutput("WaliTech API Online.");
  }

  var action = e.parameter.action;
  var result = {};

  try {
    if (action === 'getDashboardState') result = getDashboardState();
    else if (action === 'getMonthData') result = getMonthData(e.parameter.sheet);
    else if (action === 'getDetailedYearlyReport') result = getDetailedYearlyReport();
    else if (action === 'getYearlyGridStats') result = getYearlyGridStats(e.parameter.year);
    else if (action === 'getAlertSettings') result = getAlertSettings();
    else if (action === 'getAlertLogs') result = getAlertLogs();
    else if (action === 'saveAlertSettings') {
       var settings = JSON.parse(e.parameter.data);
       result = saveAlertSettings(settings);
    }
    else throw new Error("Unknown action: " + action);
    
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- API 1: DASHBOARD ---
function getDashboardState() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheetMeta = [];
    sheets.forEach(function(s) {
      if (s.getName().match(/\d{4}/) && !s.getName().includes("_")) {
        sheetMeta.push({ name: s.getName(), date: new Date("1 " + s.getName()) });
      }
    });
    sheetMeta.sort(function(a, b) { return b.date - a.date; });
    var sheetNames = sheetMeta.map(function(s) { return s.name; });
    var years = [...new Set(sheetMeta.map(s => s.date.getFullYear()))];

    var latestData = {};
    var status = "OFFLINE"; 
    var lastUpdateStr = "--";

    if (sheetNames.length > 0) {
      var currentSheetName = sheetNames[0];
      var newestSheet = ss.getSheetByName(currentSheetName);
      var lastRow = newestSheet.getLastRow();
      
      if (lastRow > 1) {
        var data = newestSheet.getRange(lastRow, 1, 1, 9).getValues()[0];
        var lastDate = new Date(data[0]);
        var now = new Date();
        
        var diffMinutes = (now - lastDate) / 1000 / 60;
        status = (diffMinutes <= 11) ? "ON-GRID" : "OFF-GRID";
        lastUpdateStr = Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

        var currEng = Number(data[7]); 
        var currCost = Number(data[8]);

        // Monthly Delta Calculation
        var prevEngRef = 0;
        var prevCostRef = 0;
        
        var summarySheet = ss.getSheetByName("Monthly_Summary");
        var foundInSummary = false;
        if (summarySheet && summarySheet.getLastRow() > 1) {
          var sumData = summarySheet.getRange(2, 1, summarySheet.getLastRow()-1, 9).getValues();
          for(var i=0; i<sumData.length; i++) {
            if(sumData[i][0] === currentSheetName) {
              prevEngRef = Number(sumData[i][3]);
              prevCostRef = Number(sumData[i][6]);
              foundInSummary = true;
              break;
            }
          }
        }

        if (!foundInSummary || prevEngRef === 0) {
           var curDateObj = new Date("1 " + currentSheetName);
           var prevDateObj = new Date(curDateObj.getFullYear(), curDateObj.getMonth() - 1, 1);
           var prevSheetName = getMonthYearString(prevDateObj);
           var prevSheet = ss.getSheetByName(prevSheetName);
           if (prevSheet) {
             var lastVals = findLastValidValues(prevSheet);
             prevEngRef = lastVals.energy;
             prevCostRef = lastVals.cost;
           }
        }

        var monthEnergy = currEng - prevEngRef;
        var monthCost = currCost - prevCostRef;
        if (monthEnergy < 0) monthEnergy = currEng; 
        if (monthCost < 0) monthCost = currCost;

        latestData = {
          ts: lastUpdateStr,
          temp: Number(data[1]), vol: Number(data[2]), cur: Number(data[3]),
          pow: Number(data[4]), freq: Number(data[5]), pf: Number(data[6]),
          yearEng: currEng, yearCost: currCost,
          monthEng: monthEnergy, monthCost: monthCost
        };
      }
    }

    var props = PropertiesService.getScriptProperties();
    var budget = props.getProperty('MONTHLY_BUDGET') || 50000;

    return { success: true, sheets: sheetNames, years: years, status: status, lastUpdate: lastUpdateStr, latest: latestData, budget: budget };

  } catch (e) { return { success: false, error: e.toString() }; }
}

// --- API 2: MONTHLY DATA (OPTIMIZED WITH DOWNSAMPLING) ---
function getMonthData(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: "Sheet '" + sheetName + "' not found" };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    // --- DYNAMIC DOWNSAMPLING ALGORITHM ---
    var MAX_POINTS = 9000; // Target maximum rows to return
    var totalDataRows = lastRow - 1; // Header is row 1
    var step = 1;

    // Calculate Step: If rows > 9000, skip data to fit fit into 9000
    if (totalDataRows > MAX_POINTS) {
      step = Math.ceil(totalDataRows / MAX_POINTS);
    }

    // Get ALL data (It is faster to fetch once than multiple times)
    var range = sheet.getRange(2, 1, totalDataRows, 9);
    var vals = range.getValues();
    
    var cleanData = [];
    
    // Extract year/month for date fallbacks
    var sheetNameParts = sheetName.split(' ');
    var sheetYear = parseInt(sheetNameParts[1], 10);
    var monthMap = {'January':0, 'February':1, 'March':2, 'April':3, 'May':4, 'June':5, 'July':6, 'August':7, 'September':8, 'October':9, 'November':10, 'December':11};
    var sheetMonth = monthMap[sheetNameParts[0]] || 0;
    var baselineDate = new Date(sheetYear, sheetMonth, 1).getTime();
    
    // --- ITERATE WITH STEP ---
    // Instead of i++, we do i += step to skip rows if needed
    for (var i = 0; i < vals.length; i += step) {
      // Ensure we don't go out of bounds (can happen on last step)
      if (!vals[i]) break;

      var row = vals[i];
      var timestamp;
      var dateValue = row[0];
      
      // Date Parsing Logic
      if (dateValue instanceof Date) {
        timestamp = dateValue.getTime();
      } else if (typeof dateValue === 'string') {
        dateValue = dateValue.replace(/\u00A0/g, ' ').trim();
        var patterns = [
          /(\d{1,2})[-/](\d{1,2})[-/](\d{4})[\sT](\d{1,2}):(\d{1,2}):(\d{1,2})/,
          /(\d{4})[-/](\d{1,2})[-/](\d{1,2})[\sT](\d{1,2}):(\d{1,2}):(\d{1,2})/,
          /(\d{1,2})\s+(\w+)\s+(\d{4})[\sT](\d{1,2}):(\d{1,2}):(\d{1,2})/
        ];
        var matched = false;
        for (var p = 0; p < patterns.length && !matched; p++) {
          var match = dateValue.match(patterns[p]);
          if (match) {
            var d, m, y, h, min, s;
            if (p === 0) { d=match[1]; m=match[2]-1; y=match[3]; }
            else if (p === 1) { y=match[1]; m=match[2]-1; d=match[3]; }
            else if (p === 2) { d=match[1]; m=['January','February','March','April','May','June','July','August','September','October','November','December'].indexOf(match[2]); y=match[3]; }
            h=match[4]||0; min=match[5]||0; s=match[6]||0;
            if(y<2000) y=sheetYear;
            timestamp = new Date(y, m, d, h, min, s).getTime();
            matched = true;
          }
        }
        if (!matched) timestamp = baselineDate + (i * 300000);
      } else if (typeof dateValue === 'number') {
        timestamp = new Date((dateValue - 25569) * 86400 * 1000).getTime();
      } else {
        timestamp = baselineDate + (i * 300000);
      }
      
      if (timestamp < 946684800000) timestamp = baselineDate + (i * 300000); // Safety check
      
      cleanData.push({
        ts: timestamp,
        tsStr: Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
        temp: isNumber(row[1]) ? Number(row[1]) : null,
        vol: isNumber(row[2]) ? Number(row[2]) : null,
        cur: isNumber(row[3]) ? Number(row[3]) : null,
        pow: isNumber(row[4]) ? Number(row[4]) : null,
        freq: isNumber(row[5]) ? Number(row[5]) : null,
        pf: isNumber(row[6]) ? Number(row[6]) : null,
        eng: isNumber(row[7]) ? Number(row[7]) : null,
        cost: isNumber(row[8]) ? Number(row[8]) : null
      });
    }
    
    return { success: true, data: cleanData };
  } catch (e) { 
    return { success: false, error: e.toString(), sheet: sheetName }; 
  }
}

function isNumber(value) { return value !== '' && value !== null && !isNaN(value) && isFinite(value); }

// --- API 3: YEARLY REPORT ---
function getDetailedYearlyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Monthly_Summary");
  if (!summarySheet || summarySheet.getLastRow() < 2) return { success: true, data: [] };
  
  var rawData = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 9).getValues();
  var structuredData = rawData.map(function(r) {
    return { year: r[1], monthIdx: r[2], energy: Number(r[5]), cost: Number(r[8]) };
  });
  return { success: true, data: structuredData };
}

// --- API 4: GRID ANALYSIS ---
function getYearlyGridStats(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var totalOn = 0, totalOff = 0;
  var stringYear = year.toString();
  
  sheets.forEach(function(s) {
    if (s.getName().indexOf(stringYear) > -1 && !s.getName().includes("_")) {
      var lastRow = s.getLastRow();
      if (lastRow > 2) {
        var times = s.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        var monthOff = 0;
        var firstTs = new Date(times[0]).getTime();
        var lastTs = new Date(times[times.length-1]).getTime();
        for(var i=1; i<times.length; i++) {
          var diff = (new Date(times[i]).getTime() - new Date(times[i-1]).getTime()) / 1000 / 60;
          if (diff > 11) monthOff += diff;
        }
        var totalSpan = (lastTs - firstTs) / 1000 / 60;
        var monthOn = Math.max(0, totalSpan - monthOff);
        totalOn += monthOn; totalOff += monthOff;
      }
    }
  });
  return { success: true, onHrs: (totalOn/60).toFixed(1), offHrs: (totalOff/60).toFixed(1) };
}

// --- API 5: SETTINGS ---
function getAlertSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    highVol: props.getProperty('HIGH_VOL') || 250,
    lowVol: props.getProperty('LOW_VOL') || 180,
    highPow: props.getProperty('HIGH_POW') || 5000,
    highTemp: props.getProperty('HIGH_TEMP') || 60,
    budget: props.getProperty('MONTHLY_BUDGET') || 50000
  };
}

function saveAlertSettings(s) { 
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    'HIGH_VOL': s.highVol, 'LOW_VOL': s.lowVol, 'HIGH_POW': s.highPow, 
    'HIGH_TEMP': s.highTemp, 'MONTHLY_BUDGET': s.budget
  }); 
  return { success: true }; 
}

function getAlertLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Alert_Logs");
  if (!sheet || sheet.getLastRow() < 2) return { success: true, logs: [] };
  var rows = Math.min(sheet.getLastRow()-1, 100);
  var vals = sheet.getRange(sheet.getLastRow()-rows+1, 1, rows, 4).getValues();
  var logs = vals.map(r => ({ time: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), type: r[1], val: r[2], msg: r[3] })).reverse();
  return { success: true, logs: logs };
}

// --- IOT RECEIVER ---
function doPost(e) {
  var res = {success: false};
  try {
    var json = JSON.parse(e.postData.contents);
    var now = new Date();
    var sheetName = getMonthYearString(now);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      initializeCleanTable(sheet);
      autoFitAllColumns(sheet);
    }
    
    var vol = parseFloat(json.voltage)||0;
    var pow = parseFloat(json.power)||0;
    var temp = parseFloat(json.temperature)||0;
    var cost = parseFloat(json.cost)||0;
    var energy = parseFloat(json.energy)||0;

    var row = [now, temp, vol, parseFloat(json.current)||0, pow, parseFloat(json.frequency)||0, parseFloat(json.powerFactor)||0, energy, cost];
    
    sheet.appendRow(row);
    formatDataRow(sheet, sheet.getLastRow());
    checkAndLogAlerts(ss, now, vol, pow, temp);
    updateMonthlySummary(ss, sheetName, energy, cost);
    
    res.success = true;
  } catch (err) { res.msg = err.toString(); }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

// --- HELPERS ---
function updateMonthlySummary(passedSS, currentSheetName, currentEnergy, currentCost) {
  if (!currentSheetName) return; 
  var ss = passedSS || SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Monthly_Summary");
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Monthly_Summary");
    summarySheet.appendRow(["Sheet_Name", "Year", "Month_Idx", "Prev_Eng_Ref", "Curr_Eng_Ref", "Actual_Energy", "Prev_Cost_Ref", "Curr_Cost_Ref", "Actual_Cost"]);
    summarySheet.setFrozenRows(1);
    summarySheet.getRange(1, 1, 1, 9).setBackground("#1e3a8a").setFontColor("white").setFontWeight("bold");
  }

  var curDate = new Date("1 " + currentSheetName);
  var prevDate = new Date(curDate.getFullYear(), curDate.getMonth() - 1, 1);
  var prevSheetName = getMonthYearString(prevDate);
  var prevEnergyEnd = 0, prevCostEnd = 0;
  
  var lastRow = summarySheet.getLastRow();
  var foundRowIndex = -1;
  
  if (lastRow > 1) {
    var finder = summarySheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(currentSheetName).matchEntireCell(true);
    var found = finder.findNext();
    if (found) foundRowIndex = found.getRow();

    var prevFinder = summarySheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(prevSheetName).matchEntireCell(true);
    var prevFound = prevFinder.findNext();
    
    if (prevFound) {
       var r = prevFound.getRow();
       var vals = summarySheet.getRange(r, 5, 1, 4).getValues()[0]; 
       prevEnergyEnd = Number(vals[0]) || 0; 
       prevCostEnd = Number(vals[3]) || 0;
    } else {
       var prevSheet = ss.getSheetByName(prevSheetName);
       if (prevSheet) {
         var v = findLastValidValues(prevSheet);
         prevEnergyEnd = v.energy;
         prevCostEnd = v.cost;
       }
    }
  } else {
     var prevSheet = ss.getSheetByName(prevSheetName);
     if (prevSheet) {
       var v = findLastValidValues(prevSheet);
       prevEnergyEnd = v.energy;
       prevCostEnd = v.cost;
     }
  }

  var actualEnergy = currentEnergy - prevEnergyEnd;
  var actualCost = currentCost - prevCostEnd;
  if (actualEnergy < 0) actualEnergy = currentEnergy; 
  if (actualCost < 0) actualCost = currentCost;

  var rowData = [currentSheetName, curDate.getFullYear(), curDate.getMonth(), prevEnergyEnd, currentEnergy, actualEnergy, prevCostEnd, currentCost, actualCost];

  if (foundRowIndex > 0) {
    summarySheet.getRange(foundRowIndex, 1, 1, 9).setValues([rowData]);
  } else {
    summarySheet.appendRow(rowData);
  }
}

function findLastValidValues(sheet) {
  if (!sheet) return { energy: 0, cost: 0 };
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { energy: 0, cost: 0 };
    var startRow = Math.max(2, lastRow - 10);
    var data = sheet.getRange(startRow, 8, lastRow - startRow + 1, 2).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var e = parseFloat(data[i][0]);
      var c = parseFloat(data[i][1]);
      if ((!isNaN(e) && e > 0) || (!isNaN(c) && c > 0)) return { energy: e || 0, cost: c || 0 };
    }
  } catch(e) {}
  return { energy: 0, cost: 0 };
}

function getMonthYearString(d) {
  var m = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return m[d.getMonth()] + " " + d.getFullYear();
}

function initializeCleanTable(sheet) {
  var headers = ["Timestamp","Temperature (°C)","Voltage (V)","Current (A)","Power (W)","Frequency (Hz)","Power Factor","Energy (kWh)","Cost (₦)"];
  sheet.getRange(1, 1, 1, 9).setValues([headers]).setBackground("#1e3a8a").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("dd/MM/yyyy HH:mm:ss"); 
  sheet.getRange("I:I").setNumberFormat("₦#,##0.00");
}

function formatDataRow(sheet, row) {
  var range = sheet.getRange(row, 1, 1, 9);
  range.setBackground(row % 2 === 0 ? "#f1f5f9" : "#ffffff").setVerticalAlignment("middle");
  sheet.getRange(row, 1).setHorizontalAlignment("left");
}

function autoFitAllColumns(sheet) {
  var w = {1: 160, 2: 100, 3: 100, 4: 100, 5: 100, 6: 100, 7: 100, 8: 120, 9: 120};
  for (var c = 1; c <= 9; c++) sheet.setColumnWidth(c, w[c]);
}