function _getEndRow(spreadsheet = SpreadsheetApp.getActiveSpreadsheet(), propServ = PropertiesService) {  // TODO: Optimize this
  console.log(propServ.getScriptProperties().getProperty("endRow"));
  console.log(getRowNum("<END>", spreadsheet) - 1);
  return getRowNum("<END>", spreadsheet) - 1;
}

function getEndRow(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
  return getRowNum("<END>", spreadsheet) - 1;
}

function testMButils(sheet = SpreadsheetApp.getActive().getActiveSheet(), storeCode = "3361") {
  console.log(SpreadsheetApp.getActiveSheet().getIndex());
  console.log(SpreadsheetApp.getActiveSpreadsheet().getSheets().length);
}

function getTotalCol() {
  return "M";
}

function getLossOverCol() {
  return String.fromCharCode(getTotalCol().charCodeAt(0) + 1);
}

function getDupFuncCol() {
  return "F";
}

function getDupLabelCol() {
  return String.fromCharCode(getDupFuncCol().charCodeAt(0) + 3);
}

function getFuncCol() {
  return "AH";
}

function getGcashButtCol() {
  return "H";
}

function getGcashButtCol2() {
  return "G";
}

function getPriceCol() {
  return "J";
}

function getLastCol() {
  return "N";
}

function getLastColIdx() {
  return 14;
}

function getProductsCol() {
  return "A";
}

function getDelCol() {
  return "C";
}

function getBegCol() {
  return "B";
}

function getEndingCol() {
  return "D";
}

function getPullOutCol() {
  return "E";
}

function getGrabCol() {
  return "I";
}

function getMBUnprotectedRangeList() {
  const startRow = 2;
  const endRow = getEndRow();
  const bsbRow = getRowNum("BSB");
  const cpbRow = getRowNum("CPB");

  return [
    'C2:E' + (bsbRow - 1),  // MB to BPB del-end-pullout
    'C' + (bsbRow + 1) + ':E' + (cpbRow - 1), // C,Patty to RSB
    'C' + (cpbRow + 1) + ':E' + endRow,
    'G2:I' + endRow,  // Order slip-fp-grab
    'A' + (endRow + 2) + ':D' + (endRow + 35), // Notes+other stocks+sheet ctrls
    getTotalCol() + (endRow + 3) + ':' + getTotalCol() + (endRow + 3),  // CoH+gcash
    getGcashButtCol() + (endRow + 4) + ":" + getGcashButtCol2() + (endRow + 4), // Gcash button
    getTotalCol() + (endRow + 8),   // panukli
    'F' + (endRow + 10) + ':' + getTotalCol() + (endRow + 37),  // Expenses
    'F' + (endRow + 40) + ':H' + (endRow + 44)  // new shift fields;
  ];
}

function updateRows() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();
  // for (x=2; x < sheets.length-3; x++) {
  for (var x = 2; x < 46; x++) {
    var sheet = sheets[x];
    console.log(x + " Updating " + sheet.getName());
    // sheet.insertRowsAfter(26,1);
    // sheet.insertRowsAfter(19,1);
    // sheet.insertRowsAfter(16,1);
    // sheet.insertRowsAfter(7,5);
    // sheet.insertColumnsAfter(10,2);
  }
}

function actualNewshift(dateObj, shiftTime, empName, propServ = PropertiesService, env = 'PRD') {
  var spreadsheet = SpreadsheetApp.getActive();
  var prevSheet = spreadsheet.getActiveSheet();
  // protectCompletedSheet(prevSheet);

  const dtFormatted = Utilities.formatDate(dateObj, "GMT+8", "MM/dd");
  const sheetName = dtFormatted + ' ' + shiftTime + ' ' + empName;

  var currentSheet = duplicateSheet(spreadsheet, sheetName);
  if (!currentSheet) return;
  //onOpen(); // Trigger function to add sheet to whitelist (auto removal of unwanted sheets)
  registerCurrentSheets(propServ);

  //spreadsheet.hideColumn(spreadsheet.getRange('C:C'));
  //spreadsheet.hideColumn(spreadsheet.getRange('E:E'));
  spreadsheet.hideColumn(spreadsheet.getRange('J:L'));
  spreadsheet.hideColumn(spreadsheet.getRange('O:S'));

  const startRow = 2;
  const endRow = getEndRow();

  // Hide salaries
  concealSalaries(true, '#ffe599', endRow, prevSheet);

  //111spreadsheet.getRange('D2:D33').copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // [COPY] Reference beginning from previous ending
  const bsbRow = getRowNum("BSB");
  const cpbRow = getRowNum("CPB");
  const prevSheetName = prevSheet.getSheetName();

  // Append year for archive sheet name referencing
  let prevSheetNameSplit = prevSheetName.split(" ");
  let today = new Date();
  let crossingNewYearAdjuster = (today.getMonth() == 0 && prevSheetNameSplit[0].startsWith["12/"]) ? 1 : 0;
  prevSheetNameSplit[0] = `${prevSheetNameSplit[0]}/${today.getFullYear() - 2000 - crossingNewYearAdjuster}`;
  const prevSheetNameYr = prevSheetNameSplit.join(" ");

  // Beginning ref to archive
  var storeCode = getStoreCodeByName(spreadsheet.getRange("A1").getValue());
  var archiveInventoryUrl = getArchiveInventoryUrl(storeCode, env);
  /*for (i = startRow; i <= endRow; i++) {
    //spreadsheet.getRange('B' + i).setFormula("'" + prevSheetName + "'!D" + i)
    var prevSheetRef = "'" + prevSheetName + "'!D" + i;
    var importRange = 'IMPORTRANGE("' + archiveInventoryUrl + '", "' + prevSheetRef + '")';
    spreadsheet.getRange('B' + i).setFormula("IFERROR(" + prevSheetRef + ", " + importRange + ")");
  }*/;
  var prevSheetRef = `'${prevSheetName}'!D${startRow}:D${endRow}`;
  var prevSheetRefYr = `'${prevSheetNameYr}'!D${startRow}:D${endRow}`;
  var importRange = 'IMPORTRANGE("' + archiveInventoryUrl + '", "' + prevSheetRefYr + '")';
  spreadsheet.getRange(`B${startRow}`).setFormula("IFERROR(IFERROR(ARRAYFORMULA(" + prevSheetRef + "), ARRAYFORMULA(" + prevSheetRefYr + ")), " + importRange + ")");

  // Panukli ref to archive
  var prevSheetRef = "'" + prevSheetName + "'!" + getTotalCol() + (endRow + 7);
  var prevSheetRefYr = "'" + prevSheetNameYr + "'!" + getTotalCol() + (endRow + 7);
  var importRange = 'IMPORTRANGE("' + archiveInventoryUrl + '", "' + prevSheetRefYr + '")';
  let secondLevelRef = "IFERROR(IFERROR(" + prevSheetRef + ", " + prevSheetRefYr + "), " + importRange + ")";
  let thirdLevelRef = "IFERROR(" + secondLevelRef + ", " + getGrabCol() + (endRow + 7) + ")";
  spreadsheet.getRange(getLossOverCol() + (endRow + 7)).setFormula(thirdLevelRef);

  // Panukli failover
  let panukliValue = prevSheet.getRange(getTotalCol() + (endRow + 7)).getValue();
  spreadsheet.getRange(getGrabCol() + (endRow + 7)).setValue(panukliValue);

  // CLEARING OPERATIONS
  spreadsheet.getRange(getTotalCol() + (endRow + 8)).setFormula("=0");  // add panukli
  spreadsheet.getRangeList([
    getTotalCol() + (endRow + 3), // CoH and Gcash
    'C2:E' + endRow, // delivery and ending
    'B3:B' + endRow, // sanity del for beg
    'G2:I' + endRow, // Tally and FP
    getPriceCol() + (endRow + 10) + ':' + getTotalCol() + (endRow + 37), // Expenses
    getLastCol() + (endRow + 10) + ':' + getLastCol() + (endRow + 30), // Expenses Salary indicator
    getLossOverCol() + (endRow + 9), // Validated button label
    'A' + (endRow + 10) + ':' + getPullOutCol() + (endRow + 10 + 18), // gcash rows
    getBegCol() + (endRow + 2) + ':' + getEndingCol() + (endRow + 8), // notes and endorsement
    getBegCol() + (endRow + 1) // message alert
  ]).clear({ contentsOnly: true, skipFilteredRows: false });

  // Patty formulas restoration (batch setFormulas)
  let pattyFormulas = [];
  for (let r = bsbRow; r <= cpbRow; r++) {
    let rowFormulas = [];
    if (r === bsbRow) {
      rowFormulas = ['C', 'D', 'E'].map(x => `=${x}${bsbRow - 3}-(${x}${bsbRow - 2}+${x}${bsbRow - 1})`);
    } else if (r === cpbRow) {
      rowFormulas = ['C', 'D', 'E'].map(x => `=${x}${cpbRow - 3}-(${x}${cpbRow - 2}+${x}${cpbRow - 1})`);
    } else {
      rowFormulas = ["", "", ""]; // Blanks for cells between BSB and CPB
    }
    pattyFormulas.push(rowFormulas);
  }
  spreadsheet.getRange(`C${bsbRow}:E${cpbRow}`).setFormulas(pattyFormulas);

  spreadsheet.getRange(getTotalCol() + (endRow + 9)).uncheck();                                                                                        // Collect button
  spreadsheet.getRange(getGcashButtCol() + (endRow + 4)).uncheck();                                                                                    // Gcash button
  spreadsheet.getRange(getLossOverCol() + (endRow + 10)).setValue(false); // add panukli

  // Message alerts
  var dayOfWeek = dateObj.getDay();
  console.log(`Checking if delivery date. Day=${dayOfWeek} must be in [2,5] and shiftTime=${shiftTime} must be PM`);
  if ([2, 5].includes(dayOfWeek) && shiftTime == "PM") {
    console.log("Setting message alert for delivery day");
    let line1 = "STOCKS ORDERING DEADLINE TODAY 10AM.";
    let msg = `${line1} \n\nLAGYAN NG BILANG ANG BUNS, SPICY CHEESE, AT LAHAT NG ITEMS. \n\nI-SEND ANG MGA KAILANGAN I-ORDER SA MESSENGER:\n  - Takip\n  - Paper bags\n  - Drinks plastic bag\n  - Straw\n  - Nachos cups\n  - Ketchup\n  - Hot sauce\n  - Tissue\n  - Garbage bag\n  - Iba pang kulang na stocks (dressing, patty, buns...)`;

    var rich = SpreadsheetApp.newRichTextValue();
    rich.setText(msg)
      .setTextStyle(0, line1.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setItalic(true)
          //.setFontFamily("Spectral")
          .setFontSize(14)
          .setForegroundColor("red")
          .build()
      )
      .setTextStyle(line1.length + 1, msg.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setItalic(true)
          //.setFontFamily("Spectral")
          .setFontSize(10)
          .setForegroundColor("red")
          .build()
      );
    spreadsheet.getRange(getBegCol() + (endRow + 1)).setRichTextValue(rich.build());
  } else if ([1, 2, 4, 6].includes(dayOfWeek) && shiftTime == "Mid") {
    console.log("Setting message for buns expiration");
    let bunColor = ((dayOfWeek == 1) ? "Blue" : (dayOfWeek == 2) ? "Green" : (dayOfWeek == 4) ? "Yellow" : "Orange");
    let line1 = `${bunColor} buns expiration today`;
    let msg = `${line1}\n\nI-report agad ang natitirang mga ${bunColor} packs ng buns sa Messenger at i-check kung mayroon nang amag o matigas na. Magsend ng pictures bago mag 10 PM.\n`;

    var rich = SpreadsheetApp.newRichTextValue();
    rich.setText(msg)
      .setTextStyle(0, line1.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setItalic(true)
          //.setFontFamily("Spectral")
          .setFontSize(14)
          .setForegroundColor("red")
          .build()
      )
      .setTextStyle(line1.length + 1, msg.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setItalic(true)
          //.setFontFamily("Spectral")
          .setFontSize(10)
          .setForegroundColor("red")
          .build()
      );
    spreadsheet.getRange(getBegCol() + (endRow + 1)).setRichTextValue(rich.build());
  }


  // Protections
  try {
    // protectDuplicatedSheet([
    //   'A1:B' + endRow, // Product & Beginning
    //   'F1:F' + (endRow+10),   // Orders + cash advance expense       F1:F43
    //   'J1:S' + (endRow+2),   // Last cols + expected sales           J1:S35
    //   getLossOverCol() + (endRow+3) + ':S' + (endRow+8),  // Loss/over  M36:S41
    //   getTotalCol() + (endRow+5) + ':' + getTotalCol() + (endRow+6),  // Expenses/spoil sum   M38:M39
    //   getFuncCol() + '3:' + getFuncCol() + '8',    // Duplicate sheet labels   W3:W8
    //   getTotalCol() + (endRow+8),       // Sales collect button   M41
    //   'A' + (endRow+32) + ":A" + (endRow+37),   // Hide functions
    //   'B' + bsbRow + ':E' + bsbRow,  // BSB
    //   'B' + cpbRow + ':E' + cpbRow  // CPB
    // ], spreadsheet)
    protectDuplicatedSheet(getMBUnprotectedRangeList(), spreadsheet, env);
  } catch (e) {
    alert(e, "MB RF Inv Err", " [Non-fatal]", env);
  }
  protectCompletedSheet(prevSheet, env);


  // Low-prio post-processing
  currentSheet.setTabColor(generateWeekDayColor(dateObj));
  fillNextShiftDetails(dateObj, shiftTime, spreadsheet, endRow);
  collectGcashToPo(endRow, prevSheet, env);
  spreadsheet.getRange(getPullOutCol() + (endRow + 2)).setValue(spreadsheet.getSheetName());
  //hideOldSheets();
  hideAllExceptRightmost(spreadsheet, true);

  // set Ready label if new sheet is next to verified
  //if (prevSheet.getRange())
  //  currentSheet.getRange(getLossOverCol() + (getEndRow()+9)).setFontColor('green').setFontStyle('italic').setFontSize(8).setValue('Verified');
  currentSheet.setCurrentCell(spreadsheet.getRange('D2'));
}

function fillNextShiftDetails(dateObj, shiftTime, spreadsheet = SpreadsheetApp.getActive(), endRow = getEndRow()) {
  spreadsheet.getRange(getDupLabelCol() + (endRow + 43) + ':' + getDupLabelCol() + (endRow + 44)).clear({ contentsOnly: true }); // Clear duplicate sheet success msg
  spreadsheet.getRange('B' + (endRow + 45)).clear({ contentsOnly: true }); // Clear duplicate sheet success msg

  var nextDate = dateObj;
  var nextShift = "PM";

  if (shiftTime == "AM") {
    nextShift = "Mid";
  } else if (shiftTime == "PM") {
    nextShift = "AM";
    nextDate.setDate(nextDate.getDate() + 1);
  }
  nextDate = Utilities.formatDate(nextDate, "GMT+8", "MM/dd");

  spreadsheet.getRange(getDupFuncCol() + (endRow + 40) + ':' + getDupFuncCol() + (endRow + 42)).setValues([[nextDate], [nextShift], [""]]);
}

function hideOldSheets(unverifiedOnly = false, endRow = getEndRow()) {
  console.log("Hiding old sheets. unverifiedOnly=" + unverifiedOnly);
  var startIndex = 2;
  // if (activeIndex != undefined) {
  //   startIndex = activeIndex-8;
  // }

  // Custom hide range
  var cellLabel = SpreadsheetApp.getActive().getActiveSheet().getRange("A" + (endRow + 32));
  var numSheets = cellLabel.getValue();
  if (isNaN(numSheets) || numSheets == 0) {
    numSheets = 3;
  }
  cellLabel.setValue("Hide old sheets:");

  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var j = startIndex; j < sheets.length - numSheets; j++) {
    // if (!sheets[i].isSheetHidden()) {
    //   if (unverifiedOnly && sheets[i].getRange('M36').isChecked()) {
    //     sheets[i].hideSheet();
    //   } else {
    //     sheets[i].hideSheet();
    //   }
    // }
    if ((!sheets[j].isSheetHidden() && !unverifiedOnly) || (!sheets[j].isSheetHidden() && sheets[j].getRange(getTotalCol() + (getEndRow(sheets[j]) + 9)).getValue() == true)) {
      console.log("Hiding sheet: " + sheets[j].getSheetName());
      concealSalaries(true, '#ffe599', endRow, sheets[j]);
      sheets[j].hideSheet();
    }
  }
}

/**
 * Hides all sheets except for the right-most sheet.
 * @param {boolean} [showLastThree=false] If true, re-shows the last 3 sheets from the right instead of just the right-most.
 */
function hideAllExceptRightmost(spreadsheet = SpreadsheetApp.getActive(), showLastThree = false) {
  var sheets = spreadsheet.getSheets();

  // Step 1: Hide all sheets except the right-most (at least one must remain visible)
  // Loop from right to left (decrementing), starting from index length-2
  for (var i = sheets.length - 2; i >= 0; i--) {
    if (!sheets[i].isSheetHidden()) {
      sheets[i].hideSheet();
    }
  }

  // Step 2: If flag is true, re-show the last 3 sheets
  if (showLastThree) {
    var start = Math.max(0, sheets.length - 3);
    // Loop from right to left (decrementing)
    for (var j = sheets.length - 1; j >= start; j--) {
      if (sheets[j].isSheetHidden()) {
        sheets[j].showSheet();
      }
    }
  }
}

function showUnverifiedSheets() {
  console.log("Showing unverified sheets");
  var sheets = SpreadsheetApp.getActive().getSheets();
  var limit = 20;
  let lastUnverifiedSheet = 0;
  for (var j = sheets.length - 1; j >= 0; j--) {
    let sheetName = sheets[j].getSheetName();
    if (sheetName == "Gcash") break;

    let endRow = getEndRow(sheets[j]);  // TODO: Can be optimized to use the getEndRow from script properties

    // Batch read M(endRow+8) and M(endRow+9) in one API call
    let sheetVals = sheets[j].getRange(getTotalCol() + (endRow + 8) + ":" + getTotalCol() + (endRow + 9)).getValues();
    let panukliVal = sheetVals[0][0];   // M(endRow+8)
    let verifiedVal = sheetVals[1][0];  // M(endRow+9)

    console.log(j + " " + sheetName + ": " + panukliVal);

    if ((sheets[j].isSheetHidden() && verifiedVal == false)) {
      sheets[j].showSheet();
      console.log("Collapsing A2:A" + (endRow + 1));
      sheets[j].getRange("A2:A" + (endRow + 1)).shiftRowGroupDepth(1).collapseGroups();
      //sheets[j].hideRows(endRow-8, 9);
      concealSalaries(false, '#000000', endRow, sheets[j]);
      lastUnverifiedSheet = j;
    } else if (j > sheets.length - 4) {
      concealSalaries(false, '#000000', endRow, sheets[j]);
    } else {
      limit--;
    }

    if (limit <= 0) {
      sheets[lastUnverifiedSheet].getRange(getLossOverCol() + (getEndRow() + 9)).setFontColor('#DDDDDD').setFontStyle('italic').setFontSize(8).setValue('Ready');
      SpreadsheetApp.flush();
      break;
    }
  }
}

function calculateGcash() {
  var spreadsheet = SpreadsheetApp.getActive();
  var gcashSheet = spreadsheet.getSheetByName('Gcash');

  var allData = gcashSheet.getRange('E2:F').getValues();
  var accumulator = 0;
  var i = 0;
  for (i = 0; i < allData.length; i++) {
    if (allData[i][0] == '') {

      break;
    }

    try {
      if (allData[i][1] == false) {
        accumulator += allData[i][0];
        gcashSheet.getRange('F' + (i + 2)).setValue(true);
      }
    } catch (e) {
      console.log(e);
    }
  }

  spreadsheet.getRange(getTotalCol() + (getEndRow() + 4)).setValue(accumulator);

  if (accumulator > 0) {
    gcashSheet.getRange('G' + (i + 1)).setValue(accumulator);
    gcashSheet.getRange('H' + (i + 1)).setValue(spreadsheet.getSheetName());
  }
}

function calculateGcashRemittance(storeName = "RF") {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = spreadsheet.getLastRow();
  const actualLastRow = lastRow;
  var range = spreadsheet.getRange("I" + lastRow);
  if (range.isBlank() || range.getValue() == "") {
    lastRow = range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
  }
  spreadsheet.getRange("I" + actualLastRow).setFormula("SUM(E" + lastRow + ":E" + actualLastRow + ")");
  addGcashToReceived(storeName, spreadsheet.getRange("I" + actualLastRow).getValue());
}

function formatGcashDateTimeColumns(e) {
  if (e.changeType == 'INSERT_ROW') {
    var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Gcash');
    spreadsheet.getRange('A:A').setNumberFormat('ddd", "mmm" "d');
    spreadsheet.getRange('B:B').setNumberFormat('h":"mm" "am/pm');
    var filtr = spreadsheet.getFilter();
    if (filtr != null) {
      filtr.remove();
    }
    spreadsheet.getRange('A:F').createFilter();
  }
}

function addSalesToCashFlow(storeName, dt, sales, gcash, expenses, cashAdvance, expectedSales, overLoss, employeeName, spoiled, dagdagPeraSaKaha, endRow, storeCode, getLockServ = null, env = 'PRD') {
  let spreadsheet = SpreadsheetApp.getActive();
  let currentSheet = spreadsheet.getActiveSheet();
  let idx = currentSheet.getIndex();
  let sheets = spreadsheet.getSheets();
  let labelRg = currentSheet.getRange(getLossOverCol() + (endRow + 9));
  let checkboxRg = currentSheet.getRange(getTotalCol() + (endRow + 9));
  let statusRg = currentSheet.getRange(getTotalCol() + (endRow + 9) + ":" + getLossOverCol() + (endRow + 9));
  let labelOrigContent = labelRg.getDisplayValue();

  // ==============================
  // PHASE 1: PRE-PROCESSING (before lock)
  // ==============================
  console.log("[PRE-PROCESS] Starting pre-processing phase...");
  checkboxRg.setFontStyle('italic').setFontSize(6).setValue("Pre-processing...");
  //SpreadsheetApp.flush();

  // Prepare cash flow row data
  let cashFlowRowData = prepareCashFlowRowData(dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha);
  console.log("[PRE-PROCESS] Cash flow row data prepared");

  // Prepare expense data (read from sheet, don't write yet)
  let expenseData = prepareExpenseData(dt, employeeName, storeCode, endRow, currentSheet, env);
  console.log("[PRE-PROCESS] Expense data prepared: " + expenseData.expenses.length + " expenses");

  // Prepare Senior/PWD data (read from sheet, don't write yet)
  let seniorData = prepareSeniorData(endRow, currentSheet, env);
  console.log("[PRE-PROCESS] Senior/PWD data prepared: " + seniorData.rowsToWrite.length + " entries");

  // Pre-load the cash flow sheet reference
  let cashFlowSheet = getCashFlowSheet(storeName, env);

  // ==============================
  // PHASE 2: LOCK ACQUISITION (3 min timeout)
  // ==============================
  console.log("[LOCK] Attempting to acquire script lock...");
  checkboxRg.setFontStyle('italic').setFontSize(6).setValue("Waiting for lock...");
  //SpreadsheetApp.flush();

  const lock = getLockServ ? getLockServ() : LockService.getScriptLock();
  const LOCK_TIMEOUT_MS = 270000; // 4.5 minutes

  let lockAcquired = false;
  let writeSuccess = false;
  try {
    lockAcquired = lock.tryLock(LOCK_TIMEOUT_MS);
  } catch (e) {
    console.error("[LOCK] Error acquiring lock: " + e.stack);
  }

  if (!lockAcquired) {
    const timeoutTime = Utilities.formatDate(new Date(), "GMT+8", "HH:mm:ss");
    console.error("[LOCK] Failed to acquire lock within timeout");
    statusRg.setValues([[false, `Timed out at ${timeoutTime} - please retry`]])
      .setFontColors([['red', 'red']])
      .setFontStyles([['normal', 'italic']])
      .setFontSizes([[12, 6]]);
    throw new Error(`Timed out waiting for lock after ${LOCK_TIMEOUT_MS / 60000} minutes at ${timeoutTime}. Please retry.`);
  }

  console.log("[LOCK] Lock acquired successfully");
  checkboxRg.setFontStyle('italic').setFontSize(6).setValue("Processing...");
  //SpreadsheetApp.flush();

  // ==============================
  // PHASE 3: WRITE OPERATIONS (inside lock)
  // ==============================
  try {
    // Write to Cash Flow
    console.log("[WRITE] Writing to Cash Flow sheet...");
    writeCashFlowRow(cashFlowSheet, cashFlowRowData);

    // Write expenses to Raw Expenses sheet
    console.log("[WRITE] Writing expenses to Raw Expenses sheet...");
    writePreparedExpenses(expenseData, env);

    // Write Senior/PWD data
    console.log("[WRITE] Writing Senior/PWD data...");
    writePreparedSeniorData(seniorData, env);

    writeSuccess = true;
  } catch (e) {
    console.error("[ERROR] Error during write phase: " + e.stack);
    statusRg.setValues([[false, "Error: " + e.message]])
      .setFontColors([['black', 'red']])
      .setFontStyles([['normal', 'italic']])
      .setFontSizes([[12, 6]]);
    throw e;
  } finally {
    // Always release the lock immediately after write operations
    console.log("[LOCK] Releasing lock...");
    lock.releaseLock();
  }

  // ==============================
  // PHASE 4: POST-WRITE (after lock release, cleanup)
  // ==============================
  if (writeSuccess) {
    try {
      console.log("[POST] Marking sheet as verified...");
      statusRg.setValues([["TRUE", 'Verified']])
        .setFontColors([['black', 'green']])
        .setFontStyles([['normal', 'italic']])
        .setFontSizes([[12, 8]]);

      if (currentSheet.getIndex() != spreadsheet.getSheets().length) {
        currentSheet.hideSheet();
      }
      currentSheet.getRange("A2").expandGroups();

      /* (With the recent optimization, there's no need to process in order)
      // Mark next inventory as ready to collect
      let sheetsLength = sheets.length;
      if (idx != sheetsLength) {
        let nextSheet = sheets[idx];
        let nextSheetLabel = nextSheet.getRange(getLossOverCol() + (getEndRow() + 9));
        let nextSheetLabelOrigVal = nextSheetLabel.getDisplayValue();
        nextSheetLabel.setFontColor('#DDDDDD').setFontStyle('italic').setFontSize(8).setValue(nextSheetLabelOrigVal + ' Ready');
      }
      */

      concealSalaries(true, '#ffe599', endRow, currentSheet);
      console.log("[POST] Sales collection completed successfully");

    } catch (e) {
      console.error("[ERROR] Error during post-write phase: " + e.stack);
      statusRg.setValues([[false, "Post-processing error: " + e.message]])
        .setFontColors([['black', 'red']])
        .setFontStyles([['normal', 'italic']])
        .setFontSizes([[12, 6]]);
      throw e;
    }
  }
}

/**
 * Prepares cash flow row data without writing to the sheet.
 * @return {Object} Contains rowValues array and salesFormula string
 */
function prepareCashFlowRowData(dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha) {
  if (!sales) {
    sales = 0;
  }

  let cashAdvanceFormula = '' + sales;
  if (cashAdvance) {
    cashAdvanceFormula = cashAdvanceFormula + '+' + cashAdvance;
  }
  if (dagdagPeraSaKaha) {
    cashAdvanceFormula = cashAdvanceFormula + '-' + dagdagPeraSaKaha;
  }

  return {
    rowValues: [[dt, sales + (cashAdvance ? '+' + cashAdvance : '') + (dagdagPeraSaKaha ? '-' + dagdagPeraSaKaha : ''), gcash, expenses, expectedSales, spoiled, overLoss, employeeName]],
    salesFormula: cashAdvanceFormula
  };
}

/**
 * Writes the prepared cash flow row data to the sheet.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cashFlowSheet The sheet to write to.
 * @param {Object} cashFlowRowData The data prepared by prepareCashFlowRowData.
 * @param {number} [appendMode=2] Strategy for finding the last row:
 *   - 0: Slowest guaranteed append
 *   - 1: Guaranteed append with a few API calls (untested)
 *   - 2: Fastest append with tendency to append extra empty rows (non-impacting)
 */
function writeCashFlowRow(cashFlowSheet, cashFlowRowData, appendMode = 2) {
  let count;
  switch (appendMode) {
    case 0:
      var colValues = cashFlowSheet.getRange("H:H").getValues();
      count = colValues.filter(String).length + 1;
      break;
    case 1:
      // Use getNextDataCell(UP) for better accuracy on specific column (H)
      // Robust check: if the last row of the sheet is filled, don't jump UP
      let maxRows = cashFlowSheet.getMaxRows();
      let lastCell = cashFlowSheet.getRange(maxRows, 8);
      if (!lastCell.isBlank()) {
        count = maxRows;
      } else {
        count = lastCell.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
        if (count === 1 && cashFlowSheet.getRange(1, 8).isBlank()) {
          count = 0;
        }
      }
      break;
    case 2:
      count = cashFlowSheet.getLastRow();
      break;
    default:
      throw new Error("Invalid append mode: " + appendMode);
  }

  cashFlowSheet.getRange(count + 1, 7, 1, 8).setValues(cashFlowRowData.rowValues);
  cashFlowSheet.getRange(count + 1, 8).setFormula(cashFlowRowData.salesFormula);
}

/**
 * Prepares expense data by reading from the inventory sheet without writing.
 * @return {Object} Contains poSpreadsheet, expenseSheetName, and expenses array
 */
function prepareExpenseData(dt, employeeName, storeCode, endRow, sheet, env = 'PRD') {
  var expenseSheetName = "Raw Expenses";
  if (storeCode == "3361") {
    expenseSheetName = expenseSheetName + " - PCGH";
  }

  let expenses = [];
  var lastRow = getEndRow(sheet);

  // Batch read expense names and amounts
  const expStartRow = lastRow + 12;
  const expNumRows = 26;
  const expNameColIdx = getDupFuncCol().charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  const expAmtColIdx = getTotalCol().charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  const expRangeVals = sheet.getRange(expStartRow, expNameColIdx, expNumRows, expAmtColIdx - expNameColIdx + 1).getValues();

  for (let idx = 0; idx < expNumRows; idx++) {
    let expenseAmt = expRangeVals[idx][expAmtColIdx - expNameColIdx];
    if (!expenseAmt) continue;
    let expenseName = expRangeVals[idx][0];
    expenses.push([dt, employeeName, expenseName, expenseAmt]);
  }

  // Senior/PWD expense
  let seniorExpenseAmt = sheet.getRange((endRow + 8), (getLastColIdx() + 1)).getValue();
  if (seniorExpenseAmt) {
    expenses.push([dt, employeeName, "Senior/PWD", seniorExpenseAmt]);
  }


  return {
    storeCode: storeCode,
    expenseSheetName: expenseSheetName,
    expenses: expenses
  };
}

/**
 * Writes prepared expense data to the Raw Expenses sheet.
 * 
 * @param {Object} expenseData The data prepared by prepareExpenseData.
 * @param {string} [env='PRD'] Environment target (PRD/UAT).
 * @param {number} [appendMode=2] Strategy for finding the last row:
 *   - 0: Slowest guaranteed append
 *   - 1: Guaranteed append with a few API calls (untested)
 *   - 2: Fastest append with tendency to append extra empty rows (non-impacting)
 */
function writePreparedExpenses(expenseData, env = 'PRD', appendMode = 2) {
  if (expenseData.expenses.length == 0) {
    console.log("[WRITE] No expenses to write");
    return;
  }

  var expenseSheetPO = getPoSpreadsheet(null, env).getSheetByName(expenseData.expenseSheetName);
  let expenseSheetPOLastRow
  switch (appendMode) {
    case 0:
      expenseSheetPOLastRow = expenseSheetPO.getRange("A:A").getValues().filter(String).length;
      break;
    case 1:
      // Use getNextDataCell(UP) for better accuracy on specific column (A)
      // Robust check: if the last row of the sheet is filled, don't jump UP
      let maxRows = expenseSheetPO.getMaxRows();
      let lastCell = expenseSheetPO.getRange(maxRows, 1);
      if (!lastCell.isBlank()) {
        expenseSheetPOLastRow = maxRows;
      } else {
        expenseSheetPOLastRow = lastCell.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
        if (expenseSheetPOLastRow === 1 && expenseSheetPO.getRange(1, 1).isBlank()) {
          expenseSheetPOLastRow = 0;
        }
      }
      break;
    case 2:
      expenseSheetPOLastRow = expenseSheetPO.getLastRow();
      break;
    default:
      throw new Error("Invalid append mode: " + appendMode);
  }

  // Batch write all expenses
  expenseSheetPO.getRange(expenseSheetPOLastRow + 1, 1, expenseData.expenses.length, 4).setValues(expenseData.expenses);

  // Copy formula for column E for each new row
  for (let i = 0; i < expenseData.expenses.length; i++) {
    expenseSheetPO.getRange("E" + expenseSheetPOLastRow).copyTo(expenseSheetPO.getRange("E" + (expenseSheetPOLastRow + 1 + i)));
  }
}

/**
 * Prepares Senior/PWD data by reading from the inventory sheet without writing.
 * @return {Object} Contains storeName, sheetName, columnIndex, and rowsToWrite array
 */
function prepareSeniorData(endRow, sheet, env = 'PRD') {
  let storeName = sheet.getRange("A1").getValue();
  let sheetName = sheet.getSheetName();

  let gcashStartRow = endRow + 10;
  let gcashEndRow = gcashStartRow + 18;

  let seniorValues = sheet.getRange("B" + gcashStartRow + ":" + "D" + gcashEndRow).getValues();
  let filteredSeniorValues = seniorValues.filter((row) => row[2] != "");

  let rowsToWrite = [];
  filteredSeniorValues.forEach((row) => {
    rowsToWrite.push([sheetName, row[0], row[2]]);
  });

  return {
    storeName: storeName,
    sheetName: sheetName,
    rowsToWrite: rowsToWrite
  };
}

/**
 * Writes prepared Senior/PWD data to the Senior/PWD sheet.
 */
function writePreparedSeniorData(seniorData, env = 'PRD', legacyImplementation = false) {
  if (seniorData.rowsToWrite.length == 0) {
    console.log("[WRITE] No Senior/PWD data to write");
    return;
  }

  let poSheet = getPoSpreadsheet(null, env).getSheetByName("Senior/PWD");

  // Find the target column for this store
  let lastColumnIndex = poSheet.getLastColumn();
  let headerValues = poSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];
  let targetColumnIndex = -1;

  for (let i = 0; i < lastColumnIndex; i++) {
    if (headerValues[i] == seniorData.storeName) {
      targetColumnIndex = i + 1; // 1-based index
      break;
    }
  }

  if (targetColumnIndex == -1) {
    throw new Error("Unable to find store code '" + seniorData.storeName + "' while writing Senior/PWD data");
  }

  let poLastRow;
  if (legacyImplementation) {
    // Detect last row for this store's column
    let columnLetter = String.fromCharCode(targetColumnIndex + 65 + 1);  // 3rd col (offset by 2)
    poLastRow = poSheet.getRange(`${columnLetter}:${columnLetter}`).getValues().filter(String).length;
  } else {
    // Detect last row for this store's column using a safe search from the bottom of the sheet
    let searchColumn = targetColumnIndex + 2;

    /* // Ensure the sheet has enough columns to avoid "Invalid argument"
    if (searchColumn > poSheet.getMaxColumns()) {
      poSheet.insertColumnsAfter(poSheet.getMaxColumns(), searchColumn - poSheet.getMaxColumns());
    } */

    // Robust check: if the last row of the sheet is filled, don't jump UP 
    let maxRows = poSheet.getMaxRows();
    let lastCell = poSheet.getRange(maxRows, searchColumn);
    if (!lastCell.isBlank()) {
      poLastRow = maxRows;
    } else {
      poLastRow = lastCell.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

      // If the search returned row 1 and row 1 is actually empty, then the column is truly empty
      if (poLastRow === 1 && poSheet.getRange(1, searchColumn).isBlank()) {
        poLastRow = 0;
      }
    }
  }

  // Batch write all Senior/PWD entries
  poSheet.getRange(poLastRow + 1, targetColumnIndex, seniorData.rowsToWrite.length, 3).setValues(seniorData.rowsToWrite);
}

function cashCollectedAppender(storeName, dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha) {
  cashCollectedAppenderWithSheetObj(getCashFlowSheet(storeName), dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha);
}

function cashCollectedAppenderWithSheetObj(cashFlowSheet, dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha) {
  let cashFlowRowData = prepareCashFlowRowData(dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha);
  writeCashFlowRow(cashFlowSheet, cashFlowRowData);
}

function addGcashToReceived(storeName, amount) {
  var cashFlowSheet = getCashFlowSheet(storeName);

  // Append value
  var colValues = cashFlowSheet.getRange("F:F").getValues();
  var count = colValues.filter(String).length + 1;
  cashFlowSheet.getRange(count + 1, 5).setValue(Utilities.formatDate(new Date(), "GMT+8", "MMM-dd"));
  cashFlowSheet.getRange(count + 1, 6).setValue(amount);
}

function extractExpenses(dt, employeeName, storeCode = "3252", endRow = getEndRow(), sheet = SpreadsheetApp.getActive().getActiveSheet(), env = 'PRD') {
  // Switcher
  var expenseSheetName = "Raw Expenses";
  if (storeCode == "3361") {
    expenseSheetName = expenseSheetName + " - PCGH";
  }
  var expenseSheetPO = getPoSpreadsheet(null, env).getSheetByName(expenseSheetName);
  var expenseSheetPOLastRow = expenseSheetPO.getRange("A:A").getValues().filter(String).length;
  console.log("Extracted expense sheet: " + sheet.getSheetName());
  var lastRow = getEndRow(sheet);

  var expenseNames = sheet.getRange(getDupFuncCol() + (lastRow + 12) + ":" + getDupFuncCol() + (lastRow + 12 + 25)).getValues();
  var expenseAmts = sheet.getRange(getTotalCol() + (lastRow + 12) + ":" + getTotalCol() + (lastRow + 12 + 25)).getValues();
  var seniorExpenseAmt = sheet.getRange((endRow + 8), (getLastColIdx() + 1)).getValue();

  var expensesToAppend = [];
  for (var i = 0; i < expenseAmts.length; i++) {
    if (expenseAmts[i][0]) {
      expensesToAppend.push([dt, employeeName, expenseNames[i][0], expenseAmts[i][0]]);
    }
  }

  if (seniorExpenseAmt) {
    expensesToAppend.push([dt, employeeName, "Senior/PWD", seniorExpenseAmt]);
  }

  if (expensesToAppend.length > 0) {
    expenseSheetPO.getRange(expenseSheetPOLastRow + 1, 1, expensesToAppend.length, 4).setValues(expensesToAppend);
    // Copy formula for column E
    expenseSheetPO.getRange("E" + expenseSheetPOLastRow).copyTo(expenseSheetPO.getRange(expenseSheetPOLastRow + 1, 5, expensesToAppend.length, 1));
  }
}

function addExpenseToRaw(sheet, expenseSheetPO, expenseSheetPOLastRow, dt, employeeName, expenseAmt = "", expenseName = "", i) {
  if (!expenseAmt) expenseAmt = sheet.getRange(getTotalCol() + i).getValue();
  if (!expenseAmt) return expenseSheetPOLastRow;
  if (!expenseName) expenseName = sheet.getRange(getDupFuncCol() + i).getValue();

  expenseSheetPO.getRange(++expenseSheetPOLastRow, 1, 1, 4).setValues([[dt, employeeName, expenseName, expenseAmt]]);
  expenseSheetPO.getRange("E" + (expenseSheetPOLastRow - 1)).copyTo(expenseSheetPO.getRange("E" + expenseSheetPOLastRow));
  return expenseSheetPOLastRow;
}

function concealSalaries(move = false, fontColor = '#ffe599', endRow = getEndRow(), sheet = SpreadsheetApp.getActive().getActiveSheet()) {
  console.log("Concealing salaries on sheet: " + sheet.getSheetName());
  var expenseCol = getDupFuncCol();
  var expenseValCol = getTotalCol();

  const startRow = endRow + 10;
  const numRows = 28;
  const expenseDataValues = sheet.getRange(expenseCol + startRow + ':' + expenseValCol + (startRow + numRows - 1)).getValues();
  const expenseFormulas = sheet.getRange(expenseValCol + startRow + ':' + expenseValCol + (startRow + numRows - 1)).getFormulas();
  const valColOffset = expenseValCol.charCodeAt(0) - expenseCol.charCodeAt(0);

  for (var idx = 0; idx < numRows; idx++) {
    var expenseName = expenseDataValues[idx][0];
    if (expenseName && expenseName.toString().toUpperCase().includes("SALARY")) {
      var salaryVal = expenseDataValues[idx][valColOffset];
      if (salaryVal) { // check if expense is populated
        console.log("Concealing: " + expenseName);
        var activeSalaryRg = sheet.getRange(expenseValCol + (startRow + idx));
        activeSalaryRg.setFontColor(fontColor);
        activeSalaryRg.offset(0, 1).setValue("<<<");

        if (move && !expenseFormulas[idx][0].startsWith("=")) {
          sheet.getRange(getPriceCol() + (startRow + idx)).setValue(salaryVal);
          activeSalaryRg.setFormula(getPriceCol() + (startRow + idx));
        }
      }
    }
  }
}

function getDelivery(storeCode = "3252", sheet = SpreadsheetApp.getActive().getActiveSheet(), env = 'PRD') {
  var poSheet = getLastPoSheet(storeCode, env);
  var poMap = constructPoMap(poSheet);

  var endRow = getEndRow();
  var items = sheet.getRange("A2:A" + endRow).getValues();
  var deliveryValues = [];

  for (var j = 0; j < items.length; j++) {
    var item = items[j][0];
    var value = null;

    if (item == "B. Patty") {
      value = poMap.get("BCB") + poMap.get("BPB") + poMap.get("BSB");
    } else if (item == "BSB" || item == "CPB") {
      deliveryValues.push([null]);
      continue;
    } else if (item == "C. Patty") {
      value = poMap.get("RSB") + poMap.get("RHB") + poMap.get("CPB");
    } else if (item.includes("powder")) {
      var powderItem = item.split(" ")[0];
      value = poMap.get(powderItem);
    } else if (item == "FT") {
      value = poMap.get(item) + poMap.get("16 OZ PAPER CUP 50'S(SORDE)");
    } else if (item == "Val Bun") {
      value = ["MB", "CB", "CT"].reduce((acc, x) => acc + (poMap.get(x) * 2), poMap.get(item));
    } else if (item == "Dbl Bun") {
      value = ["DMB", "DCB", "DCT", item].reduce((acc, x) => acc + poMap.get(x), 0);
    } else if (item == "Brio Bun") {
      value = ["BCB", "BPB", "BSB", "RSB", "CPB", "CVG", "SBR"].reduce((acc, x) => acc + (poMap.get(x) * 2), poMap.get(item)) + (poMap.get("WFC") ?? 0) + (poMap.get("WFU") ?? 0);
    } else if (item == "Htdg Bun") {
      value = ["CD", "FOF", "CCC", "BHS", item].reduce((acc, x) => acc + poMap.get(x), 0);
    } else if (item == "Premium coleslaw (BCB)") {
      value = poMap.get("BCB") / 10;
    } else if (item == "Black pepper sauce") {
      value = poMap.get("BPB") / 10;
    } else if (item == "Shawarma sauce") {
      value = poMap.get("BSB") / 10;
    } else if (item == "Veggie sauce") {
      value = poMap.get("CVG") / 10;
    } else if (item == "Veggie cabbage") {
      value = poMap.get("CVG") / 10;
    } else if (item == "Steak sauce") {
      value = poMap.get("SBR") / 10;
    } else if (item == "Steak cheese") {
      value = poMap.get("SBR") * 2;
    } else if (item == "Spicy cheese") {
      value = poMap.get("spicy") * 3;
    } else if (item == "Cheese sauce (lahat ng liquid)") {
      value = ["BCB", "BSB", "CB", "DCB"].reduce((acc, x) => acc + (poMap.get(x) / 10), (poMap.get("CCC") / 20));
    } else {
      value = poMap.get(item);
    }

    deliveryValues.push([value]);
  }

  sheet.getRange("C2:C" + endRow).setValues(deliveryValues);
}

function getPoSheets(storeCode, env = 'PRD') {
  let poSpreadsheet = getPoSpreadsheet(null, env);
  let poSheets = poSpreadsheet.getSheets();
  let filteredSheets = poSheets.filter((sheet) => sheet.getName().includes(storeCode));
  console.log("Extracted PO sheets: " + filteredSheets.map((x) => x.getSheetName()));
  return filteredSheets;
}

function getLastPoSheet(storeCode, env = 'PRD') {
  let filteredSheets = getPoSheets(storeCode, env);
  let poSheet = filteredSheets[filteredSheets.length - 1]
  console.log("Selected PO sheet: " + poSheet.getSheetName());
  return poSheet;
}

function constructPoMap(sheet) {
  var poMap = new Map();
  var startRow = 31;
  var smsRow = getRowNum("SMS NUMBERS", sheet) - 1;
  var numRows = smsRow - startRow;

  if (numRows > 0) {
    var poData = sheet.getRange(startRow, 1, numRows, 9).getValues(); // Cols A to I
    for (var idx = 0; idx < numRows; idx++) {
      var item = poData[idx][0];       // Col A
      var qty = poData[idx][8];        // Col I (index 8)
      var multiplier = poData[idx][2]; // Col C (index 2)
      var actualQty = qty * multiplier;

      if (poMap.has(item)) {
        actualQty += poMap.get(item);
        //poMap.set("spicy", parseInt((poMap.get("spicy") ?? 0) + qty)); // most likely spicy burgers
      }

      poMap.set(item, parseInt(actualQty));
    }
  }

  console.log("Constructed map size: " + poMap.size);
  return poMap;
}

function collectGcashToPo(endRow = getEndRow(), sheet = SpreadsheetApp.getActive().getActiveSheet(), env = 'PRD') {
  console.log("Collecting Gcash transactions to PO");

  let poSheet = getPoSpreadsheet(null, env).getSheetByName("GCash");
  let storeName = sheet.getRange("A1").getValue();
  //let storeCode = getStoreCodeByName(storeName)
  let sheetName = sheet.getSheetName();
  let found = false;

  for (var i = 1; i < poSheet.getLastColumn(); i++) {
    let currentCell = poSheet.getRange(1, i).getValue();
    console.log("Scanning " + currentCell);
    if (currentCell == storeName) {
      console.log("Found " + storeName);

      // Detect last row of Gcash values
      let columnLetter = String.fromCharCode(i + 65);  // 2nd col
      let poLastRow = poSheet.getRange(`${columnLetter}:${columnLetter}`).getValues().filter(String).length;

      let gcashStartRow = endRow + 10;
      let gcashEndRow = gcashStartRow + 18;
      // for (j = gcashStartRow; j < gcashEndRow; j++) {
      //   let gcashRg = getProductsCol() + j
      //   console.log("Gcash cell: " + gcashRg)
      //   let gcashVal = sheet.getRange(gcashRg).getValue();
      //   console.log("Gcash value: " + gcashVal)
      //   if (!gcashVal) break;

      //   console.log("Appending Gcash to range: " + (poLastRow+1) + "," + i)
      //   poSheet.getRange(++poLastRow, i).setValue(sheetName);
      //   poSheet.getRange(poLastRow, i+1).setValue(gcashVal);
      // }

      let gcashCol = getProductsCol();
      let gcashValues = sheet.getRange(gcashCol + gcashStartRow + ":" + gcashCol + gcashEndRow).getValues();
      let gcashValuesFiltered = gcashValues.filter((val) => val[0] > 0);

      let gcashRowsToWrite = gcashValuesFiltered.map(v => [sheetName, v[0]]);
      if (gcashRowsToWrite.length > 0) {
        console.log("Appending Gcash batch to range: " + (poLastRow + 1) + "," + i);
        poSheet.getRange(poLastRow + 1, i, gcashRowsToWrite.length, 2).setValues(gcashRowsToWrite);
      }

      found = true;
      break;
    }
  }

  if (!found) {
    throw new Error("Unable to find store code while adding Gcash transactions to PO sheet");
  }
}

function collectSeniorToPo(endRow = getEndRow(), sheet = SpreadsheetApp.getActive().getActiveSheet(), env = 'PRD') {
  console.log("Collecting Senior/PWD transactions to PO");

  let poSheet = getPoSpreadsheet(null, env).getSheetByName("Senior/PWD");
  let storeName = sheet.getRange("A1").getValue();
  //let storeCode = getStoreCodeByName(storeName)
  let sheetName = sheet.getSheetName();
  let found = false;

  // For enhancement
  //let lastCol = poSheet.getLastColumn()
  //let headerRangeVals = poSheet.getRange(1,1,1,lastCol).getValues()[0];

  let lastColumnIndex = poSheet.getLastColumn();
  let headerValues = poSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];
  for (var i = 1; i < lastColumnIndex; i++) {
    let currentCell = headerValues[i - 1];
    console.log("Scanning " + currentCell);
    if (currentCell == storeName) {
      console.log("Found " + storeName);

      // Detect last row of Gcash values
      let columnLetter = String.fromCharCode(i + 65 + 1);  // 3rd col
      let poLastRow = poSheet.getRange(`${columnLetter}:${columnLetter}`).getValues().filter(String).length;

      let gcashStartRow = endRow + 10;
      let gcashEndRow = gcashStartRow + 18;

      // // TODO: FOR ENHANCEMENT LIKE GCASH

      // for (j = gcashStartRow; j < gcashEndRow; j++) {
      //   let gcashRg = "D" + j
      //   console.log("Cell: " + gcashRg)
      //   let gcashVal = sheet.getRange(gcashRg).getValue();
      //   console.log("Value: " + gcashVal);
      //   if (!gcashVal) break;

      //   console.log("Appending Gcash to range: " + (poLastRow+1) + "," + i);
      //   poSheet.getRange(++poLastRow, i).setValue(sheetName);
      //   poSheet.getRange(poLastRow, i+1).setValue(sheet.getRange("B" + j).getValue());
      //   poSheet.getRange(poLastRow, i+2).setValue(gcashVal);
      // }

      let seniorValues = sheet.getRange("B" + gcashStartRow + ":" + "D" + gcashEndRow).getValues();
      console.log("Senior/PWD range: " + seniorValues);

      let rowsToBeWritten = seniorValues
        .filter((row) => row[2] != "")
        .map((row) => [sheetName, row[0], row[2]]);

      if (rowsToBeWritten.length > 0) {
        console.log("Appending Senior/PWD batch of " + rowsToBeWritten.length);
        poSheet.getRange(poLastRow + 1, i, rowsToBeWritten.length, 3).setValues(rowsToBeWritten);
      }

      found = true;
      break;
    }
  }

  if (!found) {
    throw new Error("Unable to find store code while adding Senior/PWD discounts to PO sheet");
  }
}

function archiveAttendance(e, env = 'PRD') {
  const ss = e.source;
  const s = ss.getActiveSheet();

  const lastRow = s.getLastRow();
  const lastCol = s.getLastColumn();
  const numRows = lastRow - 20 - 8;  // retain last 5

  const rg = s.getRange(20, 1, numRows, lastCol);
  const values = rg.getValues();
  console.log(values.length);
  console.log(numRows);

  const archiveInventory = getPoSpreadsheet(getArchiveInventoryUrl(ss.getName().split(" ")[3], env), env).getSheetByName("Attendance");
  //values.forEach(row => archiveInventory.appendRow(row))
  const archiveLastRow = archiveInventory.getLastRow();
  archiveInventory.getRange(archiveLastRow, 1, numRows, lastCol).setValues(values);
  s.deleteRows(20, lastRow - 20 - 5);
}

function verifyDelivery(rg, endRow, sheet = SpreadsheetApp.getActiveSheet(), env = 'PRD', clearFlag = false) {
  let inventoryDeliveredValue = sheet.getRange(getLossOverCol() + (endRow + 8)).getValue();
  console.log("Inventory delivered value: " + inventoryDeliveredValue);
  let storeCode = getStoreCodeByName(sheet.getRange("A1").getValue());
  let poSheets = getPoSheets(storeCode, env);

  // get inventory date
  let inventoryDate = sheet.getSheetName().split(" ", 1)[0] + "/" + (new Date().getFullYear() - 2000);
  let poSheetName = `PO D${inventoryDate} ${storeCode}`;
  console.log("Constructed PO sheet name from parsed inventory date: " + poSheetName);

  let poSheet = poSheets.find((sheet) => sheet.getSheetName() == poSheetName);
  if (poSheet == undefined) {
    rg.setFontSize(6).setFontStyle("italic").setFontColor("red").setValue(`${poSheetName} not found`);
    SpreadsheetApp.flush();
    if (clearFlag) {
      Utilities.sleep(10000);
      rg.setValue(false);
    }
    return;
  }
  console.log("Retrieved PO sheet: " + poSheet.getSheetName());

  let poVals = poSheet.getRange("O:P").getValues();
  let salesRowNum = poVals.map((row) => row[0]).indexOf("PO sales equivalent:");
  console.log("Retrieved equivalent sales row num: " + salesRowNum);
  let expectedPoSales = poVals[salesRowNum][1];
  console.log("PO equivalent value: " + expectedPoSales);

  if (inventoryDeliveredValue == expectedPoSales) {
    rg.setFontSize(6).setFontStyle("italic").setFontColor("yellow").setValue("Delivery matched PO");
  } else {
    rg.setFontSize(6).setFontStyle("italic").setFontColor("red").setValue(`Not matched: ${expectedPoSales}`);
    SpreadsheetApp.flush();
    if (clearFlag) {
      Utilities.sleep(10000);
      rg.setValue(false);
    }
  }
}

function autoFormulaEnding(spreadsheet) {
  const sheet = spreadsheet.getActiveSheet();
  const endCol = getEndingCol();
  //const delCol = getDelCol();

  for (var i = 2; i < 8; i++) {   // Adjust end iterator bound to extend beyond value burgers
    let cell = sheet.getRange(`${endCol}${i}`);
    const existingFormula = cell.getFormula();
    const leftCellAddress = cell.offset(0, -1).getA1Notation();
    let newFormula = "";

    // --- Case 1: The cell is ALREADY a formula ---
    // An empty string is falsy, a non-empty formula string is truthy.
    if (existingFormula) {
      // The existing formula already starts with '=', so we just append to it.
      newFormula = `${existingFormula}+${leftCellAddress}`;
      cell.setFormula(newFormula);
      console.log(`Appended to formula in ${cell.getA1Notation()}. New formula: "${newFormula}"`);

      // --- Case 2: The cell is NOT a formula ---
    } else {
      const originalValue = cell.getValue();

      // Check if the original value is a number before creating a new formula.
      if (typeof originalValue != 'number') {
        console.log(`Skipped: ${cell.getA1Notation()} does not contain a number and is not a formula.`);
        return;
      }

      // Construct the brand new formula from the numeric value.
      newFormula = `=${originalValue}+${leftCellAddress}`;
      cell.setFormula(newFormula);
      console.log(`Converted ${cell.getA1Notation()} to formula: "${newFormula}"`);
    }
  }
}

function pollForTimedOutCollectedSheets() {

}
