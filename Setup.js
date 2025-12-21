// --- MAINTENANCE: RUN THIS ONCE TO REBUILD SUMMARY ---
function forceRegenerateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Delete existing Monthly_Summary if it exists
  var oldSheet = ss.getSheetByName("Monthly_Summary");
  if (oldSheet) ss.deleteSheet(oldSheet);
  
  // 2. Create fresh sheet
  var summarySheet = ss.insertSheet("Monthly_Summary");
  var headers = ["Sheet_Name", "Year", "Month_Idx", "Prev_Eng_Ref", "Curr_Eng_Ref", "Actual_Energy", "Prev_Cost_Ref", "Curr_Cost_Ref", "Actual_Cost"];
  summarySheet.getRange(1, 1, 1, 9).setValues([headers])
    .setBackground("#1e3a8a").setFontColor("white").setFontWeight("bold");
  summarySheet.setFrozenRows(1);

  // 3. Get all data sheets and sort them chronologically
  var sheets = ss.getSheets();
  var sheetMeta = [];
  
  sheets.forEach(function(s) {
    // Check if sheet name looks like a date (e.g., "December 2025") and ignore system sheets
    var name = s.getName();
    if (name !== "Monthly_Summary" && name !== "Alert_Logs" && !name.includes("_")) {
      var dateObj = new Date("1 " + name);
      if (!isNaN(dateObj.getTime())) {
        sheetMeta.push({
          name: name,
          sheet: s,
          date: dateObj
        });
      }
    }
  });

  // Sort: Oldest first (to calculate running totals correctly)
  sheetMeta.sort(function(a, b) { return a.date - b.date; });

  // 4. Loop through sorted sheets and rebuild history
  var prevEngRef = 0;
  var prevCostRef = 0;

  sheetMeta.forEach(function(meta) {
    // Get last valid values from this month's sheet
    var lastVals = findLastValidValues(meta.sheet);
    var currEng = lastVals.energy;
    var currCost = lastVals.cost;

    // Calculate Delta
    var actualEnergy = currEng - prevEngRef;
    var actualCost = currCost - prevCostRef;

    // Handle resets or negative values
    if (actualEnergy < 0) actualEnergy = currEng;
    if (actualCost < 0) actualCost = currCost;

    var row = [
      meta.name,
      meta.date.getFullYear(),
      meta.date.getMonth(),
      prevEngRef,
      currEng,
      actualEnergy,
      prevCostRef,
      currCost,
      actualCost
    ];

    summarySheet.appendRow(row);

    // Update references for the next iteration
    prevEngRef = currEng;
    prevCostRef = currCost;
  });

  Logger.log("Monthly_Summary has been regenerated successfully.");
}

function sortSheetsNewestLeft() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const monthMap = {
    January: 0, February: 1, March: 2, April: 3,
    May: 4, June: 5, July: 6, August: 7,
    September: 8, October: 9, November: 10, December: 11
  };

  // Extract month + year
  const sortable = sheets.map(sheet => {
    const name = sheet.getName(); // e.g. "April 2024"
    const parts = name.split(" ");

    if (parts.length !== 2 || !(parts[0] in monthMap)) {
      return null; // Ignore non-month sheets
    }

    const month = monthMap[parts[0]];
    const year = parseInt(parts[1], 10);

    return { sheet, value: year * 12 + month };
  }).filter(Boolean);

  // Sort newest first
  sortable.sort((a, b) => b.value - a.value);

  // Move sheets
  sortable.forEach((item, index) => {
    ss.setActiveSheet(item.sheet);
    ss.moveActiveSheet(index + 1);
  });
}

// --- SIMULATION TOOL ---
function simulateIoTDevice() {
  // 1. Simulate the data your ESP32 would send
  // Change 'energy' and 'cost' slightly higher each time you run it 
  // to simulate usage over time.
  var payload = {
    "voltage": 220 + Math.random() * 10,   // Random voltage 220-230V
    "current": 6.5,
    "power": 2210,
    "frequency": 50.1,
    "powerFactor": 0.98,
    "temperature": 42.5,
    "energy": 200,  // <--- INCREASE THIS MANUALLY TO TEST MONTHLY LOGIC
    "cost": 5900   // <--- INCREASE THIS MANUALLY TO TEST MONTHLY LOGIC
  };

  // 2. Create a fake 'e' event object
  var e = {
    postData: {
      contents: JSON.stringify(payload)
    }
  };

  // 3. Run the main function
  doPost(e);
  
  Logger.log("Data simulated! Check your spreadsheet.");
}