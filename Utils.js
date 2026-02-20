function utilTest(propServ) {
  propServ.getScriptProperties().setProperty("testKey", "testVal");
}

function getRowNum(lookupVal, spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
  if (lookupVal == null) {
    throw Error("No lookupVal passed");
  }

  //console.log("Looking for '" + lookupVal + "' on sheet '" + spreadsheet.getSheetName() + "'")

  var range = spreadsheet.getRange("A:A").getValues();
  var row = 0;
  var foundFlag = false;

  for (i = row; i < range.length; i++) {
    if (range[i][0] === lookupVal) {
      row = i;
      foundFlag = true;
      break;
    }
  }

  if (!foundFlag) {
    throw Error("Unable to find lookup value '" + lookupVal + "' while searching for rownum on " + spreadsheet.getName() + ": " + spreadsheet.getSheetName());
  }

  var foundA1Notation = row + 1;
  //console.log("Found " + lookupVal + ":" + foundA1Notation)
  return foundA1Notation;
}

/**
 * Releases overall sheet protection and applies protection on passed cell ranges
 *
 * @param {string[]} protectionList - List of cell ranges to be protected
 * @param {SpreadsheetApp.Spreadsheet} [spreadsheet] - Spreadsheet to be protected
 * @return {void}
 */
function protectDuplicatedSheetLegacy(protectionCellList, spreadsheet = SpreadsheetApp.getActive(), unprotectedRanges = null) {
  try {
    spreadsheet.getActiveSheet().protect().remove();
  } catch (e) {
    console.log(e.stack);
  }

  var protection = spreadsheet.getRange('A1').protect();
  const editors = protection.getEditors();
  protection.removeEditors(editors).addEditor("mbs2edith@gmail.com");
  //protection.removeEditors(['jemnegad@gmail.com','portugal88erma@gmail.com','simarbalisado@gmail.com','sizzlingstopmain@gmail.com','jmrosin01@gmail.com','mbs2edith@gmail.com','sabandalche1995@gmail.com','kimaubreysaguinsin@gmail.com']);
  //protection.setWarningOnly(true);
  protectionCellList.forEach(function (cellRange) {
    spreadsheet.getRange(cellRange).protect().removeEditors(editors).addEditor("mbs2edith@gmail.com");
  });
  if (unprotectedRanges != null) {
    console.log("Unprotecting ranges: " + unprotectedRanges);
    protection.setUnprotectedRanges(unprotectedRanges);
  }
}

function protectDuplicatedSheet(unprotectedRangeList, spreadsheet = SpreadsheetApp.getActive()) {
  var unprotectedRanges = unprotectedRangeList.map((range) => spreadsheet.getRange(range));
  console.log("Constructed range[]: " + unprotectedRanges.map((x) => x.getA1Notation()));
  var protection = spreadsheet.getActiveSheet().protect().setDescription("Formula Protection");
  const editors = protection.getEditors();
  protection.removeEditors(editors).addEditor("mbs2edith@gmail.com");
  protection.setUnprotectedRanges(unprotectedRanges);
}

function attendance(rg) {
  let row = rg.getRow();
  if (row > 1) {
    let spreadsheet = SpreadsheetApp.getActive();
    let lastRow = spreadsheet.getLastRow();

    if (row + 2 >= lastRow) {
      spreadsheet.insertRowAfter(row);
      spreadsheet.getRange("H" + row).copyTo(spreadsheet.getRange("H" + (row + 1)));
      spreadsheet.getRange("B" + (row + 1)).insertCheckboxes();
      spreadsheet.getRange("E" + (row + 1)).insertCheckboxes();
      spreadsheet.getRange("I" + (row + 1)).insertCheckboxes();
    }
    if (rg.getA1Notation() === "B" + row && rg.isChecked() && spreadsheet.getRange("C" + row).isBlank()) {
      var dt = new Date();
      spreadsheet.getRange("C" + row).setValue(Utilities.formatDate(dt, "GMT+8", "MMM dd"));
      spreadsheet.getRange("D" + row).setValue(Utilities.formatDate(dt, "GMT+8", "HH:mm:ss"));
    } else if (rg.getA1Notation() === "E" + row && rg.isChecked() && spreadsheet.getRange("F" + row).isBlank()) {
      var dt = new Date();
      spreadsheet.getRange("F" + row).setValue(Utilities.formatDate(dt, "GMT+8", "MMM dd"));
      spreadsheet.getRange("G" + row).setValue(Utilities.formatDate(dt, "GMT+8", "HH:mm:ss"));
    }
  }
}

/**
 * Sends an email and SMS alert from an Error object
 *
 * @param {Error} Error object
 * @param {String} Email/SMS subject
 * @return {void}
 */
function alert(err, subj = "MB RF Inv Err", extraMsg = "") {
  // Send email
  // List of errors to ignore
  errorsToIgnore = [
    "Service Spreadsheets failed while accessing document with id",
    "Service error: Spreadsheets",
    'A sheet with ID "',
    "Timed out"
  ];
  if (
    ((errorMsg, errorsToIgnore) => errorsToIgnore.every((ignoreMsg) => !errorMsg.includes(ignoreMsg)))(err.stack, errorsToIgnore)
  ) {
    MailApp.sendEmail("bakulinglings@gmail.com", subj, err.stack + extraMsg);
  }

  // Send SMS
  var smsSheet = getPoSpreadsheet("https://docs.google.com/spreadsheets/d/17yPemlid9FVMdzVDX8Eg8Tu1W-zOg_prNtQeUeEidAg/edit").getSheetByName("SMS");
  var smsLastRow = smsSheet.getLastRow();
  var range = smsSheet.getRange("A" + smsLastRow);
  if (range.isBlank() || range.getValue() === "") {
    smsLastRow = range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
  }
  //smsSheet.getRange("A" + smsLastRow).setValue("+639763715943")
  smsSheet.getRange("A" + smsLastRow).setValue("+639151272800");
  smsSheet.getRange("B" + smsLastRow).setValue(subj + ": " + err.toString() + extraMsg);
  smsSheet.getRange("C" + smsLastRow).insertCheckboxes();
  smsSheet.getRange("C" + smsLastRow).setValue(true);
}

function protectCompletedSheet(sheet = SpreadsheetApp.getActive().getActiveSheet()) {
  try {
    sheet.protect().remove();
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var sheetProtect = sheet.protect().setDescription("Prevent changes on past inventories");
    sheetProtect = sheetProtect.removeEditors(sheetProtect.getEditors());
    sheetProtect.addEditor("mbs2edith@gmail.com");

    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.canEdit()) {
        console.log("removing " + protection.getRange());
        protection.remove();
      }
    }
  } catch (e) {
    console.log(e.stack);
  }
}

function showLastNSheets(numSheets = 10) {
  console.log("Showing last " + numSheets + " sheets");
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (i = sheets.length - 1; i >= sheets.length - numSheets; i--) {
    sheets[i].showSheet();
  }
}

function deleteSheet() {
  SpreadsheetApp.getActive().deleteActiveSheet();
}

function activateLastSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();
  spreadsheet.setActiveSheet(sheets[sheets.length - 1]);
}

function registerCurrentSheets(propServ = PropertiesService) {
  propServ.getScriptProperties().setProperty("sheetName", JSON.stringify(SpreadsheetApp.getActiveSpreadsheet().getSheets().map((s) => s.getSheetName())));
}

function deleteUnregisteredSheets(e, propServ = PropertiesService, lockServ = LockService) {
  const lock = lockServ.getDocumentLock();
  if (lock.tryLock(350000)) {
    try {
      if (e.changeType != "INSERT_GRID") return;
      const sheetNames = JSON.parse(propServ.getScriptProperties().getProperty("sheetName"));
      const dumpSite = getPoSpreadsheet("https://docs.google.com/spreadsheets/d/1rPiSSJlbfLDKjqofRxqh1DR1lP5VYS67SFBp99to5Iw/edit");
      e.source.getSheets().forEach((s) => {
        var sheetName = s.getSheetName();
        if (!sheetNames.includes(sheetName)) {
          console.log("Moving to dumpsite: " + sheetName);
          e.source.getSheetByName(sheetName).copyTo(dumpSite);
          e.source.deleteSheet(s);
        }
      });
    } catch (e) {
      throw new Error(JSON.stringify(e));
    } finally {
      lock.releaseLock();
    }
  } else {
    throw new Error("timeout");
  }
}

function generateWeekDayColor(dt) {
  const colors = [
    '#FF0000',  // red
    '#FF9900',  // orange
    '#FFFF00',  // yellow
    '#00FF00',  // green
    '#00FFFF',  // cyan
    '#0000FF',  // blue
    '#FF00FF',  // magenta
  ];
  var day = dt.getDay();
  console.log("Generating color for date: " + dt + " = " + day);

  return colors[day];
}

function generateColor(str) {
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  var colour = '#';
  for (var i = 0; i < 3; i++) {
    var value = (hash >> (i * 8)) & 0xFF;
    colour += ('00' + value.toString(16)).substr(-2);
  }
  console.log(colour);
  return colour;
}

function duplicateSheet(spreadsheet, sheetName) {
  try {
    spreadsheet.duplicateActiveSheet();
    var currentSheet = spreadsheet.getActiveSheet();
    currentSheet.setName(sheetName);
  } catch (e) {
    if (e.message.includes("already exists")) {
      console.log("Dumping '" + currentSheet.getSheetName() + "' due to error: " + e.stack);
      const dumpSite = getPoSpreadsheet("https://docs.google.com/spreadsheets/d/1rPiSSJlbfLDKjqofRxqh1DR1lP5VYS67SFBp99to5Iw/edit");
      currentSheet.copyTo(dumpSite);
      spreadsheet.deleteSheet(currentSheet);
    }
    return false;
  }
  return currentSheet;
}

function incrementLeftCell(sheet = SpreadsheetApp.getActive().getActiveSheet(), cell) {
  try {
    var origRange = sheet.getRange(cell);
    var row = origRange.getRow();
    var col = origRange.getColumn();
    console.log(row + " " + col);
    console.log(Number(row) + " " + (col - 1));
    var leftCellRange = sheet.getRange(Number(row), (col - 1));
    var origVal = leftCellRange.getValue();
    leftCellRange.setValue(origVal + 1);
  } catch (e) {
    console.error(e.stack);
  }
}

function getStoreCodeByName(storeName) {
  if (storeName == "RF") {
    return "3252";
  } else if (storeName == "PCGH") {
    return "3361";
  } else {
    throw Error("Error getting store code. Invalid store name: " + storeName);
  }
}

function getStoreCodes() {
  return ["3252", "3361"];
}

function getCashFlowSheet(storeName) {
  // Store switch
  var cashFlowSheetName = "Cash flow";
  if (storeName != "RF") {
    cashFlowSheetName = "Cash flow - " + storeName;
  }
  console.log("Constructing cash flow sheet name: " + cashFlowSheetName);

  var cashFlowSheet = getPoSpreadsheet().getSheetByName(cashFlowSheetName);

  console.log("Acquired cash flow sheet: " + cashFlowSheet.getSheetName());

  return cashFlowSheet;
}

function getInventoryUrl(storeCode) {
  console.log("Store code: " + storeCode);
  var urlMap = new Map();
  urlMap.set("3252", 'https://docs.google.com/spreadsheets/d/1XxOw-t7q60ULv59GzABMn4uEouFZlZrYtMOMsoLtAtA/edit');
  urlMap.set("3361", 'https://docs.google.com/spreadsheets/d/1rwW-JantrQTuEg2Uzez6XR0xcvceOFbMVuke1t2Idak/edit');

  var url = urlMap.get(storeCode + "");
  console.log("Acquired URL: " + url);
  if (!url) {
    throw new Error("Unable to lookup URL for store code: " + storeCode);
  } else {
    return url;
  }

  var inventoryUrl = 'https://docs.google.com/spreadsheets/d/1XxOw-t7q60ULv59GzABMn4uEouFZlZrYtMOMsoLtAtA/edit';
  if (storeCode == "3361") {
    console.log("Switching inventoryUrl");
    inventoryUrl = 'https://docs.google.com/spreadsheets/d/1rwW-JantrQTuEg2Uzez6XR0xcvceOFbMVuke1t2Idak/edit';
  }
  return inventoryUrl;
}

function getArchiveInventoryUrl(storeCode) {
  console.log("Store code: " + storeCode);
  var inventoryUrl = 'https://docs.google.com/spreadsheets/d/1-31Kf3SsdMhNu1D9ziwBvqlhI6LUTlMgO3PD-vUTNkM/edit';
  if (storeCode == "3361") {
    console.log("Switching inventoryUrl");
    inventoryUrl = 'https://docs.google.com/spreadsheets/d/1QasIoOag67V1I_ePVT_-LaENPTA0Ah1cPIu7uo-q7XE/edit';
  }
  return inventoryUrl;
}

function getPoUrl() {
  return 'https://docs.google.com/spreadsheets/d/10IGAlAy_4LqyFgi3UNAVHBDy9hp69yJOd6oQVhL4JLk/edit';
}

function getPoSpreadsheet(url = getPoUrl()) {
  while (true) {
    try {
      return SpreadsheetApp.openByUrl(url);
      break;
    } catch (e) {
      console.error(e.stack);
      Utilities.sleep(5000);
    }
  }
}

function removeConflictedSheets() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheets = spreadsheet.getSheets();
  let sheetNames = sheets.map((sheet) => sheet.getSheetName());
  let conflictSheetNames = sheetNames.filter((sheetName) => sheetName.includes("conflict") || sheetName.startsWith("Copy of ") || sheetName.startsWith("Kopya ng ") || sheetName.startsWith("Sheet"));
  conflictSheetNames.forEach((sheetName) => {
    console.log(`Deleting conflicted sheet: ${sheetName}`);
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetName));
  });
}

function triggerFuncWithProcessingText(rangeA1Not, func, spreadsheet = SpreadsheetApp.getActive()) {
  let msgRg = spreadsheet.getRange(rangeA1Not);
  let existingMsg = msgRg.getValue();
  if (existingMsg == "Processing...") existingMsg = "";
  msgRg.setValue("Processing...");
  SpreadsheetApp.flush();
  try {
    func();
  } catch (err) {
    msgRg.setValue(err.stack);
    SpreadsheetApp.flush();
    Utilities.sleep(10000);
    msgRg.setValue(existingMsg);

    throw err;
  }
  msgRg.setValue(existingMsg);
}

function hideSheets(ss = SpreadsheetApp.getActive()) {
  ss.getSheets().slice(2).filter((sheet) => !sheet.isSheetHidden()).forEach((sheet) => {
    console.log("Hiding sheet: " + sheet.getSheetName());
    sheet.hideSheet();
  });
}

function shuffle(array) {
  let currentIndex = array.length, randomIndex;

  // While there remain elements to shuffle.
  while (currentIndex > 0) {

    // Pick a remaining element.
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex--;

    // And swap it with the current element.
    [array[currentIndex], array[randomIndex]] = [
      array[randomIndex], array[currentIndex]];
  }

  return array;
}

function deleteAllSheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();

  if (sheets.length > 2) {
    for (i = 2; i < sheets.length; i++) {
      console.log("Deleting sheet: " + sheets[i].getSheetName());
      spreadsheet.deleteSheet(sheets[i]);
    }
  }
}

function endsWithSpaceAndNumber(str) {
  // Regex:
  // \s  - matches a whitespace character (space, tab, newline, etc.)
  // \d+ - matches one or more digits
  // $   - asserts position at the end of the string
  const regex = /\s\d+$/;
  return regex.test(str);
}

function cleanupCopies(customStartIndex = -50, customEndIndex = -1) {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().slice(customStartIndex, customEndIndex).filter((s) => s.getSheetName().startsWith("Copy of ") || s.getSheetName().startsWith("Kopya ng ") || endsWithSpaceAndNumber(s.getSheetName()) || (s.getSheetName().startsWith("Sheet") && s.getSheetName() != "Sheet1")).forEach((s) => {
    if (s.getSheetName().startsWith("Sheet")) {
      console.log(`Deleting sheet: ${s.getSheetName()}`);
      ss.deleteSheet(s);
    } else {
      try {
        if (s.getSheetName().startsWith("Copy of ")) {
          s.setName(s.getSheetName().substring(8));
        } else {
          s.setName(s.getSheetName().substring(9));
        }
      } catch (e) {
        if (e.stack.includes("already exists")) {
          console.log(`Deleting redundant sheet: ${s.getSheetName()}`);
          ss.deleteSheet(s);
        }
      }
    }
  });
}

function rasterizeSheets(customRgFormulaCheck = "B2", customStartIndex = -50, customEndIndex = -1) {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().slice(customStartIndex, customEndIndex).filter((s) => s.getRange(customRgFormulaCheck).getFormula().substring(0, 1) == "=").forEach((s) => {
    console.log(`Rasterizing ${s.getSheetName()}`);
    const rg = s.getRange(1, 1, s.getLastRow(), s.getLastColumn());
    rg.setDataValidation(null);
    rg.setValues(rg.getValues());
  });
}

function checkAndReMergeRanges(e, sheet, rangeList) {
  // Optional: You could filter by changeType, but sometimes row/column changes
  // implicitly unmerge, so checking regardless might be safer.
  if (e.changeType !== 'FORMAT' && e.changeType !== 'OTHER') { // Example filter
    return;
  }

  //const sheet = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log(`Change detected on sheet: ${sheet.getName()}. Checking merge status for ranges: ${rangeList.join(', ')}`);

  rangeList.forEach((rangeString) => {
    try {
      const range = sheet.getRange(rangeString);

      // Check if the *intended* range is currently NOT merged
      if (!range.isMerged()) {
        range.merge();  // ignore the rest for now. Force re-merge
        return;
        // Check if *any* cell within the range is part of *some* merged range
        // This helps avoid errors if only *part* of the intended merge was broken
        // or if the range definition is slightly off but overlaps a valid merge.
        // A simple heuristic: check the top-left cell.
        const topLeftCell = range.getCell(1, 1);
        if (!topLeftCell.isPartOfMerge()) {
          Logger.log(`Range ${rangeString} was unmerged. Re-merging.`);
          // If it's definitely unmerged, re-merge it.
          range.merge();
        } else {
          // It's possible topLeftCell is part of a *different* merge now.
          // Or maybe the unmerge affected other parts of the range.
          // Advanced logic could go here to check the entire range boundary.
          // For now, we log this potentially ambiguous state.
          Logger.log(`Range ${rangeString} is not fully merged as defined, but top-left cell (${topLeftCell.getA1Notation()}) is part of *a* merge. No action taken to avoid potential conflicts.`);
        }
      } else {
        // Logger.log(`Range ${rangeString} is correctly merged.`); // Optional: uncomment for verbose logging
      }
    } catch (error) {
      // Log an error if the range string is invalid (e.g., due to deleted rows/cols)
      Logger.log(`Error processing range "${rangeString}": ${error}`);
      throw error;
    }
  });
}

/**
 * Searches a given range and returns the Range object for the first cell
 * that is formatted as a checkbox.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to search.
 * @return {GoogleAppsScript.Spreadsheet.Range | null} The Range object for the first checkbox found, or null if none are found.
 */
function findFirstCheckboxInRange(range) {
  const validations = range.getDataValidations();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const sheet = range.getSheet();

  // Iterate through the 2D array of data validation rules.
  for (let r = 0; r < validations.length; r++) {
    for (let c = 0; c < validations[r].length; c++) {

      const rule = validations[r][c];

      // Check if the cell has a rule and if that rule's type is CHECKBOX.
      if (rule != null && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {

        // We found the first one. Return the specific cell's Range object and stop searching.
        return sheet.getRange(startRow + r, startCol + c);
      }
    }
  }

  // If the loops complete, no checkbox was found in the range.
  return null;
}