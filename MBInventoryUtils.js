function _getEndRow(spreadsheet=SpreadsheetApp.getActiveSpreadsheet(), propServ=PropertiesService) {
  console.log(propServ.getScriptProperties().getProperty("endRow"))
  console.log(getRowNum("<END>", spreadsheet)-1)
  return getRowNum("<END>", spreadsheet)-1;
}

function getEndRow(spreadsheet=SpreadsheetApp.getActiveSpreadsheet()) {
  return getRowNum("<END>", spreadsheet)-1;
}

function testMButils(sheet=SpreadsheetApp.getActive().getActiveSheet(), storeCode="3361") {
  console.log(SpreadsheetApp.getActiveSheet().getIndex())
  console.log(SpreadsheetApp.getActiveSpreadsheet().getSheets().length)
}

function getTotalCol() {
  return "M"
}

function getLossOverCol() {
  return String.fromCharCode(getTotalCol().charCodeAt(0)+1);
}

function getDupFuncCol() {
  return "F";
}

function getDupLabelCol() {
  return String.fromCharCode(getDupFuncCol().charCodeAt(0)+3);
}

function getFuncCol() {
  return "AH";
}

function getGcashButtCol(){
  return "H";
}

function getGcashButtCol2(){
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
  return "D"
}

function getPullOutCol() {
  return "E"
}

function getGrabCol() {
  return "I"
}

function getMBUnprotectedRangeList() {
  const startRow = 2
  const endRow = getEndRow()
  const bsbRow = getRowNum("BSB");
  const cpbRow = getRowNum("CPB")

  return [
    'C2:E' + (bsbRow-1),  // MB to BPB del-end-pullout
    'C' + (bsbRow+1) + ':E' + (cpbRow-1), // C,Patty to RSB
    'C' + (cpbRow+1) + ':E' + endRow,
    'G2:I' + endRow,  // Order slip-fp-grab
    'A' + (endRow+2) + ':D' + (endRow+35), // Notes+other stocks+sheet ctrls
    getTotalCol() + (endRow+3) + ':' + getTotalCol() + (endRow+3),  // CoH+gcash
    getGcashButtCol() + (endRow+4) + ":" + getGcashButtCol2() + (endRow+4), // Gcash button
    getTotalCol() + (endRow+8),   // panukli
    'F' + (endRow+10) + ':' + getTotalCol() + (endRow+37),  // Expenses
    'F' + (endRow+40) + ':H' + (endRow+44)  // new shift fields
  ]
}

function updateRows() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();
  // for (x=2; x < sheets.length-3; x++) {
  for (x=2; x < 46; x++) {
    var sheet = sheets[x];
    console.log(x + " Updating " + sheet.getName());
    // sheet.insertRowsAfter(26,1)
    // sheet.insertRowsAfter(19,1)
    // sheet.insertRowsAfter(16,1)
    // sheet.insertRowsAfter(7,5)
    // sheet.insertColumnsAfter(10,2)
  }
}

function actualNewshift(dateObj, shiftTime, empName, propServ=PropertiesService) {
  var spreadsheet = SpreadsheetApp.getActive();
  var prevSheet = spreadsheet.getActiveSheet();  
  // protectCompletedSheet(prevSheet);

  const dtFormatted = Utilities.formatDate(dateObj, "GMT+8", "MM/dd");
  const sheetName = dtFormatted + ' ' + shiftTime + ' ' + empName;
  
  var currentSheet = duplicateSheet(spreadsheet, sheetName)
  if(!currentSheet) return;
  //onOpen(); // Trigger function to add sheet to whitelist (auto removal of unwanted sheets)
  registerCurrentSheets(propServ)
  
  //spreadsheet.hideColumn(spreadsheet.getRange('C:C'))
  //spreadsheet.hideColumn(spreadsheet.getRange('E:E'))
  spreadsheet.hideColumn(spreadsheet.getRange('J:L'))
  spreadsheet.hideColumn(spreadsheet.getRange('O:S'))

  const startRow = 2
  const endRow = getEndRow()

  // Hide salaries
  concealSalaries(true, '#ffe599', endRow, prevSheet);

  //111spreadsheet.getRange('D2:D33').copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // [COPY] Reference beginning from previous ending
  const bsbRow = getRowNum("BSB");
  const cpbRow = getRowNum("CPB")
  const prevSheetName = prevSheet.getSheetName()

  // Append year for archive sheet name referencing
  let prevSheetNameSplit = prevSheetName.split(" ")
  let today = new Date()
  let crossingNewYearAdjuster = (today.getMonth() === 0 && prevSheetNameSplit[0].startsWith["12/"]) ? 1 : 0 ;
  prevSheetNameSplit[0] = `${prevSheetNameSplit[0]}/${today.getFullYear()-2000-crossingNewYearAdjuster}`
  const prevSheetNameYr = prevSheetNameSplit.join(" ") 

  // Beginning ref to archive
  var storeCode = getStoreCodeByName(spreadsheet.getRange("A1").getValue());
  var archiveInventoryUrl = getArchiveInventoryUrl(storeCode)
  /*for (i = startRow; i <= endRow; i++) {
    //spreadsheet.getRange('B' + i).setFormula("'" + prevSheetName + "'!D" + i)
    var prevSheetRef = "'" + prevSheetName + "'!D" + i
    var importRange = 'IMPORTRANGE("' + archiveInventoryUrl + '", "' + prevSheetRef + '")';
    spreadsheet.getRange('B' + i).setFormula("IFERROR(" + prevSheetRef + ", " + importRange + ")"); 
  }*/
  var prevSheetRef = `'${prevSheetName}'!D${startRow}:D${endRow}`
  var prevSheetRefYr = `'${prevSheetNameYr}'!D${startRow}:D${endRow}`
  var importRange = 'IMPORTRANGE("' + archiveInventoryUrl + '", "' + prevSheetRefYr + '")';
  spreadsheet.getRange(`B${startRow}`).setFormula("IFERROR(IFERROR(ARRAYFORMULA(" + prevSheetRef + "), ARRAYFORMULA(" + prevSheetRefYr + ")), " + importRange + ")");
  
  // Panukli ref to archive
  var prevSheetRef = "'" + prevSheetName + "'!" + getTotalCol() + (endRow+7)
  var prevSheetRefYr = "'" + prevSheetNameYr + "'!" + getTotalCol() + (endRow+7)
  var importRange = 'IMPORTRANGE("' + archiveInventoryUrl + '", "' + prevSheetRefYr + '")';
  let secondLevelRef = "IFERROR(IFERROR(" + prevSheetRef + ", " + prevSheetRefYr + "), " + importRange + ")"
  let thirdLevelRef = "IFERROR(" + secondLevelRef + ", " + getGrabCol() + (endRow+7) + ")"
  spreadsheet.getRange(getLossOverCol() + (endRow+7)).setFormula(thirdLevelRef);

  // Panukli failover
  let panukliValue = prevSheet.getRange(getTotalCol() + (endRow+7)).getValue()
  spreadsheet.getRange(getGrabCol() + (endRow+7)).setValue(panukliValue)

  // CLEARING OPERATIONS  
  spreadsheet.getRange(getTotalCol() + (endRow+8)).setFormula("=0")  // add panukli
  spreadsheet.getRange(getTotalCol() + (endRow+3) + ':' + getTotalCol() + (endRow+3)).clear({contentsOnly: true, skipFilteredRows: true});// CoH and Gcash
  spreadsheet.getRange('C2:E' + endRow).clear({contentsOnly: true, skipFilteredRows: true});                // delivery and ending
  spreadsheet.getRange('B3:B' + endRow).clear({contentsOnly: true, skipFilteredRows: false});                // sanity del for beg
  // Patty formulas
  (['C', 'D', 'E']).forEach(function(x, i) {
    spreadsheet.getRange(x + bsbRow).setFormula(x + (bsbRow-3) + '-(' + x + (bsbRow-2) + "+" + x +(bsbRow-1) + ')')
    spreadsheet.getRange(x + cpbRow).setFormula(x + (cpbRow-3) + '-(' + x + (cpbRow-2) + "+" + x + (cpbRow-1) + ')')
  })
  spreadsheet.getRange('G2:I' + endRow).clear({contentsOnly: true, skipFilteredRows: true});                                              // Tally and FP
  spreadsheet.getRange(getPriceCol() + (endRow+10) + ':' + getTotalCol() + (endRow+37)).clear({contentsOnly: true, skipFilteredRows: true}); // Expenses
  spreadsheet.getRange(getLastCol() + (endRow+10) + ':' + getLastCol() + (endRow+30)).clear({contentsOnly: true, skipFilteredRows: true}); // Expenses Salary indicator
  spreadsheet.getRange(getTotalCol() + (endRow+9)).uncheck();                                                                                        // Collect button
  spreadsheet.getRange(getGcashButtCol() + (endRow+4)).uncheck();                                                                                    // Gcash button
  spreadsheet.getRange(getLossOverCol() + (endRow+9)).clear({contentsOnly: true, skipFilteredRows: true});    // Validated button label
  //spreadsheet.getRange(getTotalCol() + (endRow+8)).clear({contentsOnly: true, skipFilteredRows: true});    // Add panukli
  spreadsheet.getRange("A" + (endRow+10) + ":" + getPullOutCol() + (endRow+10+18)).clear({contentsOnly: true, skipFilteredRows: false});    // gcash rows
  spreadsheet.getRange(getBegCol() + (endRow+1)).setValue("") // message alert
  spreadsheet.getRange(getLossOverCol() + (endRow+10)).setValue(false) // add panukli
  spreadsheet.getRange(getBegCol() + (endRow+2) + ":" + getEndingCol() + (endRow+8)).clear({contentsOnly: true, skipFilteredRows: true}) // notes and endorsement
  
  // Message alerts
  var dayOfWeek = dateObj.getDay();
  console.log(`Checking if delivery date. Day=${dayOfWeek} must be in [2,5] and shiftTime=${shiftTime} must be PM`)
  if([2,5].includes(dayOfWeek) && shiftTime == "PM") {
    console.log("Setting message alert for delivery day")
    let line1 = "STOCKS ORDERING DEADLINE TODAY 10AM."
    let msg = `${line1} 

LAGYAN NG BILANG ANG BUNS, SPICY CHEESE, AT LAHAT NG ITEMS. 

I-SEND ANG MGA KAILANGAN I-ORDER SA MESSENGER:
  - Takip
  - Paper bags
  - Drinks plastic bag
  - Straw
  - Nachos cups
  - Ketchup
  - Hot sauce
  - Tissue
  - Garbage bag
  - Iba pang kulang na stocks (dressing, patty, buns...)`

    var rich = SpreadsheetApp.newRichTextValue()
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
      .setTextStyle(line1.length+1, msg.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setItalic(true)
          //.setFontFamily("Spectral")
          .setFontSize(10)
          .setForegroundColor("red")
          .build()
    )
    spreadsheet.getRange(getBegCol() + (endRow+1)).setRichTextValue(rich.build())
  } else if ([1,2,4,6].includes(dayOfWeek) && shiftTime == "Mid") {
    console.log("Setting message for buns expiration")
    let bunColor = ((dayOfWeek == 1) ? "Blue" : (dayOfWeek == 2) ? "Green" : (dayOfWeek == 4) ? "Yellow" : "Orange")
    let line1 = `${bunColor} buns expiration today`
    let msg = `${line1}

I-report agad ang natitirang mga ${bunColor} packs ng buns sa Messenger at i-check kung mayroon nang amag o matigas na. Magsend ng pictures bago mag 10 PM.
`

    var rich = SpreadsheetApp.newRichTextValue()
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
      .setTextStyle(line1.length+1, msg.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setItalic(true)
          //.setFontFamily("Spectral")
          .setFontSize(10)
          .setForegroundColor("red")
          .build()
    )
    spreadsheet.getRange(getBegCol() + (endRow+1)).setRichTextValue(rich.build())
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
    protectDuplicatedSheet(getMBUnprotectedRangeList(), spreadsheet)
  } catch (e) {
    alert(e, "MB RF Inv Err", " [Non-fatal]")
  }
  protectCompletedSheet(prevSheet);


  // Low-prio post-processing
  currentSheet.setTabColor(generateWeekDayColor(dateObj));
  fillNextShiftDetails(dateObj, shiftTime, spreadsheet, endRow)
  collectGcashToPo(endRow, prevSheet);
  spreadsheet.getRange(getPullOutCol() + (endRow+2)).setValue(spreadsheet.getSheetName());
  hideOldSheets();

  // set Ready label if new sheet is next to verified
  //if (prevSheet.getRange())
  //  currentSheet.getRange(getLossOverCol() + (getEndRow()+9)).setFontColor('green').setFontStyle('italic').setFontSize(8).setValue('Verified')
  currentSheet.setCurrentCell(spreadsheet.getRange('D2'));
}

function fillNextShiftDetails(dateObj, shiftTime, spreadsheet=SpreadsheetApp.getActive(), endRow=getEndRow()) {
  spreadsheet.getRange(getDupLabelCol() + (endRow+43) + ':' + getDupLabelCol() + (endRow+44)).clear({contentsOnly: true}) // Clear duplicate sheet success msg
  spreadsheet.getRange('B' + (endRow+45)).clear({contentsOnly: true}) // Clear duplicate sheet success msg

  var nextDate = dateObj;
  var nextShift = "PM";
  
  if (shiftTime==="AM") {
    nextShift = "Mid"
  } else if (shiftTime==="PM") {
    nextShift = "AM"
    nextDate.setDate(nextDate.getDate() + 1)
  }
  nextDate = Utilities.formatDate(nextDate, "GMT+8", "MM/dd");

  spreadsheet.getRange(getDupFuncCol() + (endRow+40) + ':' + getDupFuncCol() + (endRow+42)).setValues([[nextDate],[nextShift],[""]])
}

function hideOldSheets(unverifiedOnly=false, endRow=getEndRow()) {
  console.log("Hiding old sheets. unverifiedOnly=" + unverifiedOnly)
  var startIndex = 2
  // if (activeIndex !== undefined) {
  //   startIndex = activeIndex-8
  // }

  // Custom hide range
  var cellLabel = SpreadsheetApp.getActive().getActiveSheet().getRange("A" + (endRow+32))
  var numSheets = cellLabel.getValue()
  if (isNaN(numSheets) || numSheets == 0) {
    numSheets = 3
  }
  cellLabel.setValue("Hide old sheets:")

  var sheets = SpreadsheetApp.getActive().getSheets()
  for (j = startIndex; j < sheets.length-numSheets; j++) {
    // if (!sheets[i].isSheetHidden()) {
    //   if (unverifiedOnly && sheets[i].getRange('M36').isChecked()) {
    //     sheets[i].hideSheet()
    //   } else {
    //     sheets[i].hideSheet()
    //   }
    // }
    if ((!sheets[j].isSheetHidden() && !unverifiedOnly) || (!sheets[j].isSheetHidden() && sheets[j].getRange(getTotalCol() + (getEndRow(sheets[j])+9)).getValue() === true)) {
      console.log("Hiding sheet: " + sheets[j].getSheetName())
      concealSalaries(true, '#ffe599', endRow, sheets[j])
      sheets[j].hideSheet()
    }
  }
}

function showUnverifiedSheets() {
  console.log("Showing unverified sheets")
  var sheets = SpreadsheetApp.getActive().getSheets()
  var limit = 20;
  let lastUnverifiedSheet = 0
  for (j = sheets.length-1; j >= 0; j--) {
    if (sheets[j].getSheetName() == "Gcash") break;

    let endRow = getEndRow(sheets[j])

    console.log(j + " " + sheets[j].getSheetName() + ": " + sheets[j].getRange(getTotalCol() + (endRow+8)).getValue())

    if ((sheets[j].isSheetHidden() && sheets[j].getRange(getTotalCol() + (endRow+9)).getValue() === false)) {
      sheets[j].showSheet();
      console.log("Collapsing A2:A" + (endRow+1))
      sheets[j].getRange("A2:A" + (endRow+1)).shiftRowGroupDepth(1).collapseGroups()
      //sheets[j].hideRows(endRow-8, 9)
      concealSalaries(false, '#000000', endRow, sheets[j])
      lastUnverifiedSheet = j
    } else if (j > sheets.length-4) {
      concealSalaries(false, '#000000', endRow, sheets[j])
    } else {
      limit--;
    }

    if (limit <= 0) {
      sheets[lastUnverifiedSheet].getRange(getLossOverCol() + (getEndRow()+9)).setFontColor('#DDDDDD').setFontStyle('italic').setFontSize(8).setValue('Ready')
      break;
    }
  }
}

function calculateGcash() {
  var spreadsheet = SpreadsheetApp.getActive()
  var gcashSheet = spreadsheet.getSheetByName('Gcash')

  var allData = gcashSheet.getRange('E2:F').getValues()
  var accumulator = 0;
  var i = 0
  for (i = 0; i < allData.length; i++) {
    if (allData[i][0] === '') {

      break;
    }
    
    try {
      if (allData[i][1] === false) {
        accumulator += allData[i][0];
        gcashSheet.getRange('F' + (i+2)).setValue(true)
      }
    } catch (e) {
      console.log(e)
    }
  }

  spreadsheet.getRange(getTotalCol() + (getEndRow()+4)).setValue(accumulator)

  if (accumulator > 0) {
    gcashSheet.getRange('G' + (i+1)).setValue(accumulator)
    gcashSheet.getRange('H' + (i+1)).setValue(spreadsheet.getSheetName())
  }
}

function calculateGcashRemittance(storeName="RF") {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = spreadsheet.getLastRow();
  const actualLastRow = lastRow;
  var range = spreadsheet.getRange("I" + lastRow);
  if (range.isBlank() || range.getValue() === "") {
    lastRow = range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow()+1;
  }
  spreadsheet.getRange("I" + actualLastRow).setFormula("SUM(E" + lastRow + ":E" + actualLastRow + ")")
  addGcashToReceived(storeName, spreadsheet.getRange("I" + actualLastRow).getValue())
}

function formatGcashDateTimeColumns(e) {
  if (e.changeType === 'INSERT_ROW') {
    var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Gcash');
    spreadsheet.getRange('A:A').setNumberFormat('ddd", "mmm" "d');
    spreadsheet.getRange('B:B').setNumberFormat('h":"mm" "am/pm');
    var filtr = spreadsheet.getFilter();
    if(filtr != null) {
      filtr.remove();
    }
    spreadsheet.getRange('A:F').createFilter();
  }
}

function addSalesToCashFlow(storeName, dt, sales, gcash, expenses, cashAdvance, expectedSales, overLoss, employeeName, spoiled, dagdagPeraSaKaha, endRow) {
  let spreadsheet = SpreadsheetApp.getActive()
  let currentSheet = spreadsheet.getActiveSheet()
  let cashFlowSheet = getCashFlowSheet(storeName)

  // Hold until previous sheet has completed processing
  let idx = currentSheet.getIndex();
  let sheets = spreadsheet.getSheets();
  if (idx > 3) {  // Skip for first sheet (including attendance and gcash)
    let previousSheet = sheets[idx-2]
    console.log("Previous sheet name: " + previousSheet.getSheetName())
    let previousLabelRg = previousSheet.getRange(getLossOverCol() + (getEndRow()+9))
    let previousRg = previousSheet.getRange(getTotalCol() + (getEndRow()+9))
    let labelRg = currentSheet.getRange(getLossOverCol() + (getEndRow()+9))
    let labelOrigContent = ""
    let isLabelDisplayed = false

    let timeoutInterval = 15   // 3 minutes timeout (10 sec interval) - reduced from 18 to 10 due to not accurate sleep
    //while (previousRg.getValue() !== "Verified") {
    while (!String(previousLabelRg.getValue()).includes("Timed out") && previousRg.getValue() === "Processing..." && timeoutInterval-- > 0) {
      console.log("Current value: " + previousRg.getValue())
      if (!isLabelDisplayed) {
        labelOrigContent = labelRg.getDisplayValue()
        labelRg.setFontStyle('italic').setFontSize(6).setValue(labelOrigContent + " Waiting for previous sheet to complete")
        isLabelDisplayed = true
      }
      Utilities.sleep(10000)
      SpreadsheetApp.flush()
    }

    if (String(previousLabelRg.getValue()).includes("Timed out") ||timeoutInterval <= 0) {
      labelRg.setValue("Timed out")
      let currentRg = currentSheet.getRange(getTotalCol() + (getEndRow()+9))
      currentRg.setValue("")
      currentRg.setValue(false)
      throw new Error("Timed out")
    }

    if (isLabelDisplayed) labelRg.setValue(labelOrigContent)
  }
  // End of pre-check

  // Add to Cash Collected
  cashCollectedAppenderWithSheetObj(cashFlowSheet, dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha)

  // expenses
  extractExpenses(dt, employeeName, storeCode, endRow, currentSheet)

  collectSeniorToPo(endRow, currentSheet)
  currentSheet.getRange(getTotalCol() + (endRow+9)).setValue("TRUE")
  currentSheet.getRange(getLossOverCol() + (getEndRow()+9)).setFontColor('green').setFontStyle('italic').setFontSize(8).setValue('Verified')
  if (currentSheet.getIndex() != spreadsheet.getSheets().length) {
    currentSheet.hideSheet()
  }
  currentSheet.getRange("A2").expandGroups()
  // Optional step to mark next inventory as ready to collect
  let sheetsLength = sheets.length;

  if(idx != sheetsLength) { // Not the last sheet
    let nextSheet = sheets[idx] // no need to increment as getIndex() is 1-based
    let nextSheetLabel = nextSheet.getRange(getLossOverCol() + (getEndRow()+9))
    let nextSheetLabelOrigVal = nextSheetLabel.getDisplayValue()
    nextSheetLabel.setFontColor('#DDDDDD').setFontStyle('italic').setFontSize(8).setValue(nextSheetLabelOrigVal + ' Ready')
  }

  concealSalaries(true, '#ffe599', endRow, currentSheet)
}

function cashCollectedAppender(storeName, dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha) {
    cashCollectedAppenderWithSheetObj(getCashFlowSheet(storeName), dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha)
}

function cashCollectedAppenderWithSheetObj(cashFlowSheet, dt, sales, cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName, dagdagPeraSaKaha) {
    // Check if remittance is blank
    if (!sales) {
      sales = 0;
    }

    // Append value
    var colValues = cashFlowSheet.getRange("H:H").getValues();
    var count = colValues.filter(String).length + 1
    if(cashAdvance) {
      cashAdvance = '+' + cashAdvance;
    }
    if(dagdagPeraSaKaha) {
      cashAdvance = cashAdvance + '-' + dagdagPeraSaKaha
    }
      // old implem
      /*cashFlowSheet.getRange(count+1,7).setValue(dt);
      cashFlowSheet.getRange(count+1,8).setFormula(sales + cashAdvance);
      cashFlowSheet.getRange(count+1,9).setValue(gcash);
      cashFlowSheet.getRange(count+1,10).setValue(expenses);
      cashFlowSheet.getRange(count+1,11).setValue(expectedSales);
      cashFlowSheet.getRange(count+1,12).setValue(spoiled);
      cashFlowSheet.getRange(count+1,13).setValue(overLoss);
      cashFlowSheet.getRange(count+1,14).setValue(employeeName);*/

    // new implem
    let rowToBeWritten = [[dt, sales + cashAdvance, gcash, expenses, expectedSales, spoiled, overLoss, employeeName]]
    cashFlowSheet.getRange(count+1,7,1,8).setValues(rowToBeWritten)
    cashFlowSheet.getRange(count+1,8).setFormula(sales + cashAdvance)
}

function addGcashToReceived(storeName, amount) {
  var cashFlowSheet = getCashFlowSheet(storeName)

  // Append value
  var colValues = cashFlowSheet.getRange("F:F").getValues();
  var count = colValues.filter(String).length + 1
  cashFlowSheet.getRange(count+1,5).setValue(Utilities.formatDate(new Date(), "GMT+8", "MMM-dd"));
  cashFlowSheet.getRange(count+1,6).setValue(amount);
}

function extractExpenses(dt, employeeName, storeCode="3252", endRow=getEndRow(), sheet=SpreadsheetApp.getActive().getActiveSheet()) {
  // Switcher
  var expenseSheetName = "Raw Expenses"
  if(storeCode == "3361") {
    expenseSheetName = expenseSheetName + " - PCGH"
  }
  var expenseSheetPO = getPoSpreadsheet().getSheetByName(expenseSheetName);
  var expenseSheetPOLastRow = expenseSheetPO.getRange("A:A").getValues().filter(String).length
  console.log("Extracted expense sheet: " + sheet.getSheetName())
  var lastRow = getEndRow(sheet);

  for(i = (lastRow+12); i < (lastRow+12+26); i++) {
    expenseSheetPOLastRow = addExpenseToRaw(sheet, expenseSheetPO, expenseSheetPOLastRow, dt, employeeName)
  }

  // Senior/PWD
  const x = addExpenseToRaw(sheet, expenseSheetPO, expenseSheetPOLastRow, dt, employeeName, sheet.getRange((endRow+8), (getLastColIdx()+1)).getValue(), "Senior/PWD")
}

function addExpenseToRaw(sheet, expenseSheetPO, expenseSheetPOLastRow, dt, employeeName, expenseAmt="", expenseName="") {
  if (!expenseAmt) expenseAmt = sheet.getRange(getTotalCol() + i).getValue();
  if (!expenseAmt) return expenseSheetPOLastRow;
  if (!expenseName) expenseName = sheet.getRange(getDupFuncCol() + i).getValue();
  
  // expenseSheetPO.getRange("A" + ++expenseSheetPOLastRow).setValue(dt)
  // expenseSheetPO.getRange("B" + expenseSheetPOLastRow).setValue(employeeName)
  // expenseSheetPO.getRange("C" + expenseSheetPOLastRow).setValue(expenseName)
  // expenseSheetPO.getRange("D" + expenseSheetPOLastRow).setValue(expenseAmt)
  expenseSheetPO.getRange(++expenseSheetPOLastRow, 1, 1, 4).setValues([[dt, employeeName, expenseName, expenseAmt]])
  expenseSheetPO.getRange("E" + (expenseSheetPOLastRow-1)).copyTo(expenseSheetPO.getRange("E" + expenseSheetPOLastRow))
  return expenseSheetPOLastRow;
}

function concealSalaries(move=false, fontColor='#ffe599', endRow=getEndRow(), sheet=SpreadsheetApp.getActive().getActiveSheet()) {
  console.log("Concealing salaries on sheet: " + sheet.getSheetName())
  var expenseCol = getDupFuncCol();
  var expenseValCol = getTotalCol();

  //let expenseColIdx = expenseCol.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  //let expenseValColIdx = expenseValCol.charCodeAt(0) - 'A'.charCodeAt(0) + 1;

  //let expenseNames = sheet.getRange((endRow+10), expenseColIdx, (endRow+10+28), 1).getValues()[0]
  //let expenseVals = sheet.getRange((endRow+10), expenseValCol, (endRow+10+28), 1).getValues()[0]
  for (i=(endRow+10); i<(endRow+10+28); i++) {
    var expenseName = sheet.getRange(expenseCol + i).getValue();
    //var expenseName = expenseNames[i-1]
    
    if (expenseName.toString().toUpperCase().includes("SALARY")) {
      var activeSalaryRg = sheet.getRange(expenseValCol + i)

      if(activeSalaryRg.getValue()) { // check if expense is populated
        console.log("Concealing: " + expenseName)
        activeSalaryRg.setFontColor(fontColor);
        activeSalaryRg.offset(0,1).setValue("<<<")

        if (move && !activeSalaryRg.getFormula().startsWith("=")) {
          //activeSalaryRg.copyTo(sheet.getRange(getPriceCol() + i))
          sheet.getRange(getPriceCol() + i).setValue(activeSalaryRg.getValue())
          activeSalaryRg.setFormula(getPriceCol() + i)
          
          // if (activeSalaryRg.getFormula() != ("=" + getPriceCol() + i)) { // check first if value has already been moved/hidden
          //   activeSalaryRg.copyTo(sheet.getRange(getPriceCol() + i))
          //   activeSalaryRg.setFormula(getPriceCol() + i)
          // }
        }
      }
    }
  }
}

function getDelivery(storeCode="3252",  sheet=SpreadsheetApp.getActive().getActiveSheet()) {
  var poSheet = getLastPoSheet(storeCode)
  var poMap = constructPoMap(poSheet)

  for (var i = 2; i <= getEndRow(); i++) {
    var item = sheet.getRange("A" + i).getValue();
    //console.log("[DEBUG] Current item: " + i + "=" + item)
    var value = null;

    if (item === "B. Patty") {
      value = poMap.get("BCB") + poMap.get("BPB") + poMap.get("BSB")
    } else if (item === "BSB" || item === "CPB") {
      continue;
    } else if (item === "C. Patty") {
      value = poMap.get("RSB") + poMap.get("RHB") + poMap.get("CPB")
    } else if (item.includes("powder")) {
      /*if(item === "CLT powder") {
        value = poMap.get("CLT")/5
      } else if (item === "FT powder") {
        value = poMap.get("FT")/5
      } else {
        value = poMap.get(item.split(" ")[0])
      }*/
      var powderItem = item.split(" ")[0]
      value = poMap.get(powderItem)
      if (["FT", "CLT"].includes(powderItem)) {
        value = value   // Previously value = value/5
      }
    } else if (item === "FT") {
      value = poMap.get(item) + poMap.get("16 OZ PAPER CUP 50'S(SORDE)")
    } else if (item === "Val Bun") {
      value = ["MB", "CB", "CT"].reduce((acc,x) => acc + (poMap.get(x) * 2), poMap.get(item))
      //value = poMap.get(item) + ((poMap.get("MB") + poMap.get("CB") + poMap.get("CT")) * 2)
    } else if (item === "Dbl Bun") {
      //value = poMap.get(item) + poMap.get("DMB") + poMap.get("DCB") + poMap.get("DCT")
      value = ["DMB", "DCB", "DCT", item].reduce((acc,x) => acc + poMap.get(x), 0)
    } else if (item === "Brio Bun") {
      value = ["BCB", "BPB", "BSB", "RSB", "CPB", "CVG", "SBR"].reduce((acc,x) => acc + (poMap.get(x) * 2), poMap.get(item)) + (poMap.get("WFC") ?? 0) + (poMap.get("WFU") ?? 0)
      //value = poMap.get(item) + ((poMap.get("BCB") + poMap.get("BPB") + poMap.get("BSB") + poMap.get("RSB") + poMap.get("CPB") + poMap.get("CVG") + poMap.get("SBR")) * 2) + (poMap.get("WFC") ?? 0) + (poMap.get("WFU") ?? 0)
    } else if (item === "Htdg Bun") {
      value = ["CD", "FOF", "CCC", "BHS", item].reduce((acc,x) => acc + poMap.get(x), 0)
      //value = poMap.get(item) + poMap.get("CD") + poMap.get("FOF") + poMap.get("CCC")
    } else if (item === "Premium coleslaw (BCB)") {
      value = poMap.get("BCB")/10
    } else if (item === "Black pepper sauce") {
      value = poMap.get("BPB")/10
    } else if (item === "Shawarma sauce") {
      value = poMap.get("BSB")/10
    } else if (item === "Veggie sauce") {
      value = poMap.get("CVG")/10
    } else if (item === "Veggie cabbage") {
      value = poMap.get("CVG")/10
    } else if (item === "Steak sauce") {
      value = poMap.get("SBR")/10
    } else if (item === "Steak cheese") {
      value = poMap.get("SBR")*2
    } else if (item === "Spicy cheese") {
      value = poMap.get("spicy")*3
    } else if (item === "Cheese sauce (lahat ng liquid)") {
      //value = poMap.get("BCB")/10 + poMap.get("BSB")/10 + poMap.get("CCC")/20 + poMap.get("CB")/10 + poMap.get("DCB")/10
      value = ["BCB", "BSB", "CB", "DCB"].reduce((acc,x) => acc + (poMap.get(x) / 10), (poMap.get("CCC") / 20))
    } else {
      value = poMap.get(item);
    }

    sheet.getRange("C" + i).setValue(value)
  }
}

function getPoSheets(storeCode) {
  let poSpreadsheet = getPoSpreadsheet();
  let poSheets = poSpreadsheet.getSheets();
  let filteredSheets = poSheets.filter((sheet) => sheet.getName().includes(storeCode))
  console.log("Extracted PO sheets: " + filteredSheets.map((x) => x.getSheetName()))
  return filteredSheets
}

function getLastPoSheet(storeCode) {
  let filteredSheets = getPoSheets(storeCode)
  let poSheet = filteredSheets[filteredSheets.length - 1]
  console.log("Selected PO sheet: " + poSheet.getSheetName())
  return poSheet;
}

function constructPoMap(sheet) {
  var poMap = new Map()
  for(var i = 31; i < getRowNum("SMS NUMBERS",sheet)-1; i++) {
    var item = sheet.getRange("A" + i).getValue();
    var qty = sheet.getRange("I" + i).getValue();
    var multiplier = sheet.getRange("C" + i).getValue();
    var actualQty = qty*multiplier

    if (poMap.has(item)) {
      actualQty += poMap.get(item)
      poMap.set("spicy", parseInt((poMap.get("spicy") ?? 0) + qty)) // most likely spicy burgers
    }

    poMap.set(item, parseInt(actualQty))
  }

  console.log("Constructed map size: " + poMap.size)
  return poMap;
}

function collectGcashToPo(endRow=getEndRow(), sheet=SpreadsheetApp.getActive().getActiveSheet()) {
  console.log("Collecting Gcash transactions to PO")

  let poSheet = getPoSpreadsheet().getSheetByName("GCash")
  let storeName = sheet.getRange("A1").getValue()
  //let storeCode = getStoreCodeByName(storeName)
  let sheetName = sheet.getSheetName()
  let found = false
  
  for (i=1; i < poSheet.getLastColumn(); i++) {
    let currentCell = poSheet.getRange(1, i).getValue();
    console.log("Scanning " + currentCell)
    if (currentCell == storeName) {
      console.log("Found " + storeName)

      // Detect last row of Gcash values
      let columnLetter = String.fromCharCode(i+65)  // 2nd col
      let poLastRow = poSheet.getRange(`${columnLetter}:${columnLetter}`).getValues().filter(String).length

      let gcashStartRow = endRow + 10
      let gcashEndRow = gcashStartRow + 18
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
      //console.log("GCash values: " + gcashValues)
      let gcashValuesFiltered = gcashValues.filter((val) => val > 0)
      //console.log("Filtered GCash values: " + gcashValuesFiltered.length)
      gcashValuesFiltered.forEach((gcashVal) => {
        console.log("Appending Gcash to range: " + (poLastRow+1) + "," + i)
        poSheet.getRange(++poLastRow, i).setValue(sheetName);
        poSheet.getRange(poLastRow, i+1).setValue(gcashVal);
      })

      found = true
      break
    }
  }

  if (!found) {
    throw new Error("Unable to find store code while adding Gcash transactions to PO sheet")
  }
}

function collectSeniorToPo(endRow=getEndRow(), sheet=SpreadsheetApp.getActive().getActiveSheet()) {
  console.log("Collecting Senior/PWD transactions to PO")

  let poSheet = getPoSpreadsheet().getSheetByName("Senior/PWD")
  let storeName = sheet.getRange("A1").getValue()
  //let storeCode = getStoreCodeByName(storeName)
  let sheetName = sheet.getSheetName()
  let found = false
  
  // For enhancement
  //let lastCol = poSheet.getLastColumn()
  //let headerRangeVals = poSheet.getRange(1,1,1,lastCol).getValues()[0];

  let lastColumnIndex = poSheet.getLastColumn()
  let headerValues = poSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0]
  for (i=1; i < lastColumnIndex; i++) {
    let currentCell = headerValues[i-1];
    console.log("Scanning " + currentCell)
    if (currentCell == storeName) {
      console.log("Found " + storeName)

      // Detect last row of Gcash values
      let columnLetter = String.fromCharCode(i+65+1)  // 3rd col
      let poLastRow = poSheet.getRange(`${columnLetter}:${columnLetter}`).getValues().filter(String).length

      let gcashStartRow = endRow + 10
      let gcashEndRow = gcashStartRow + 18

      // // TODO: FOR ENHANCEMENT LIKE GCASH

      // for (j = gcashStartRow; j < gcashEndRow; j++) {
      //   let gcashRg = "D" + j
      //   console.log("Cell: " + gcashRg)
      //   let gcashVal = sheet.getRange(gcashRg).getValue();
      //   console.log("Value: " + gcashVal)
      //   if (!gcashVal) break;
        
      //   console.log("Appending Gcash to range: " + (poLastRow+1) + "," + i)
      //   poSheet.getRange(++poLastRow, i).setValue(sheetName);
      //   poSheet.getRange(poLastRow, i+1).setValue(sheet.getRange("B" + j).getValue());
      //   poSheet.getRange(poLastRow, i+2).setValue(gcashVal);
      // }

      let seniorValues = sheet.getRange("B" + gcashStartRow + ":" + "D" + gcashEndRow).getValues();
      console.log("Senior/PWD range: " + seniorValues)
      
      let filteredSeniorValues = seniorValues.filter((row) => row[2] != "")
      console.log("Filtered Senior/PWD: " + filteredSeniorValues)

      let rowsToBeWritten = []
      filteredSeniorValues.forEach((row) => {
        console.log("Appending to Senior/PWD: " + row[0] + " = " + row[2])
        // poSheet.getRange(++poLastRow, i).setValue(sheetName);
        // poSheet.getRange(poLastRow, i+1).setValue(row[0]);
        // poSheet.getRange(poLastRow, i+2).setValue(row[2]);
        rowsToBeWritten.push([sheetName, row[0], row[2]])
      })
      if(filteredSeniorValues.length > 0)
        poSheet.getRange(++poLastRow, i, filteredSeniorValues.length, 3).setValues(rowsToBeWritten)

      found = true
      break
    }
  }

  if (!found) {
    throw new Error("Unable to find store code while adding Senior/PWD discounts to PO sheet")
  }
}

function archiveAttendance(e) {
  const ss = e.source
  const s = ss.getActiveSheet();

  const lastRow = s.getLastRow();
  const lastCol = s.getLastColumn();
  const numRows = lastRow-20-8;  // retain last 5

  const rg = s.getRange(20, 1, numRows, lastCol)
  const values = rg.getValues();
  console.log(values.length)
  console.log(numRows)
  
  const archiveInventory = getPoSpreadsheet(getArchiveInventoryUrl(ss.getName().split(" ")[3])).getSheetByName("Attendance")
  //values.forEach(row => archiveInventory.appendRow(row))
  const archiveLastRow = archiveInventory.getLastRow()
  archiveInventory.getRange(archiveLastRow, 1, numRows, lastCol).setValues(values)

  s.deleteRows(20, lastRow-20-5);
}

function verifyDelivery(rg, endRow, sheet=SpreadsheetApp.getActiveSheet()) {
  let inventoryDeliveredValue = sheet.getRange(getLossOverCol() + (endRow+8)).getValue()
  console.log("Inventory delivered value: " + inventoryDeliveredValue)
  let storeCode = getStoreCodeByName(sheet.getRange("A1").getValue())
  let poSheets = getPoSheets(storeCode)

  // get inventory date
  let inventoryDate = sheet.getSheetName().split(" ", 1)[0] + "/" + (new Date().getFullYear() - 2000)
  let poSheetName = `PO D${inventoryDate} ${storeCode}`
  console.log("Constructed PO sheet name from parsed inventory date: " + poSheetName)

  let poSheet = poSheets.find(sheet => sheet.getSheetName() === poSheetName)
  if (poSheet == undefined) {
    rg.setFontSize(6).setFontStyle("italic").setFontColor("red").setValue(`${poSheetName} not found`)
    SpreadsheetApp.flush()
    Utilities.sleep(10000)
    rg.setValue(false)
    return;
  }
  console.log("Retrieved PO sheet: " + poSheet.getSheetName())

  let poVals = poSheet.getRange("O:P").getValues()
  let salesRowNum = poVals.map(row => row[0]).indexOf("PO sales equivalent:")
  console.log("Retrieved equivalent sales row num: " + salesRowNum)
  let expectedPoSales = poVals[salesRowNum][1]
  console.log("PO equivalent value: " + expectedPoSales)

  if (inventoryDeliveredValue == expectedPoSales) {
    rg.setFontSize(6).setFontStyle("italic").setFontColor("yellow").setValue("Delivery matched PO")
  } else {
    rg.setFontSize(6).setFontStyle("italic").setFontColor("red").setValue(`Not matched: ${expectedPoSales}`)
    SpreadsheetApp.flush()
    Utilities.sleep(10000)
    rg.setValue(false)
  }
}

function autoFormulaEnding(spreadsheet) {
  const sheet = spreadsheet.getActiveSheet()
  const endCol = getEndingCol()
  //const delCol = getDelCol()

  for (i=2; i<8; i++) {   // Adjust end iterator bound to extend beyond value burgers
    let cell = sheet.getRange(`${endCol}${i}`)
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
      if (typeof originalValue !== 'number') {
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