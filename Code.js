function test(s = SpreadsheetApp.getActiveSheet()) {
  let ss = SpreadsheetApp.getActive();

  ss.getSheets().slice(-8).forEach((x) => {
    console.log(x.getSheetName());
  });
  return;
  spreadsheet.getRange("N71").setValue("TRUE");

}

function adhocTrigger() {
  manualTrigger("F104");
}

function manualTrigger(rangeStr) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  SpreadsheetApp.setActiveSheet(sheets[sheets.length - 1]);
  const e = {
    source: SpreadsheetApp.getActive(),
    range: SpreadsheetApp.getActiveSheet().getRange(rangeStr),
    isChecked: true,
  };
  installedOnEditTrigger(e);
}

function onOpen() {
  registerCurrentSheets();
}

function installedOnChange(e) {
  deleteUnregisteredSheets(e);
}

function installedOnEditTrigger(e, propServ = PropertiesService, env = 'PRD') {
  let rg = e.range;
  // console.log(`[DEBUG] Acquired range from event: ${JSON.stringify(rg, null, 2)}`);
  // // Check if more than one cell was edited to determine the right strategy.
  // if (rg.getNumRows() > 1 || rg.getNumColumns() > 1) {
  //   // For multi-cell edits, search the entire range for the first checkbox.
  //   rg = findFirstCheckboxInRange(rg);
  // }
  // console.log(`[DEBUG] Filtered checkbox range: ${JSON.stringify(rg, null, 2)}`);

  const endRow = parseInt(propServ.getScriptProperties().getProperty("endRow")) || getEndRow();
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const a1Not = rg.getA1Notation();
  try {
    if (rg.isChecked()) {
      rg.uncheck();

      /*if (SpreadsheetApp.getActive().getSheetName() == "Attendance") {    // Insert checkboxes on attendance sheet
        if (a1Not == "H1") {
          triggerFuncWithProcessingText(a1Not, function() {archiveAttendance(e)}, spreadsheet);
        } else {
          attendance(rg);
        }
        return;

      } else*/ if (a1Not == getDupFuncCol() + (endRow + 43)) {
        console.log("Starting new shift script");

        // Cleanup
        spreadsheet.getRange(getDupFuncCol() + (endRow + 40) + ':' + getDupFuncCol() + (endRow + 42)).setBorder(false, false, false, false, false, false, null, SpreadsheetApp.BorderStyle.SOLID);
        spreadsheet.getRange(getDupLabelCol() + (endRow + 40) + ':' + getDupLabelCol() + (endRow + 44)).clear({ contentsOnly: true, skipFilteredRows: true });

        var dt = spreadsheet.getRange(getDupFuncCol() + (endRow + 40));
        var shft = spreadsheet.getRange(getDupFuncCol() + (endRow + 41));
        var nam = spreadsheet.getRange(getDupFuncCol() + (endRow + 42));

        if (dt.isBlank()) {
          dt.setBorder(true, true, true, true, false, false, "red", SpreadsheetApp.BorderStyle.DASHED);
          spreadsheet.getRange(getDupLabelCol() + (endRow + 40)).setValue('<- Set a date').setFontColor("red").setFontWeight("bold");
          return;
        }
        if (shft.isBlank()) {
          shft.setBorder(true, true, true, true, false, false, "red", SpreadsheetApp.BorderStyle.DASHED);
          spreadsheet.getRange(getDupLabelCol() + (endRow + 41)).setValue('<- Set AM or PM').setFontColor("red").setFontWeight("bold");
          return;
        }
        if (nam.isBlank()) {
          nam.setBorder(true, true, true, true, false, false, "red", SpreadsheetApp.BorderStyle.DASHED);
          spreadsheet.getRange(getDupLabelCol() + (endRow + 42)).setValue('<- Set your name').setFontColor("red").setFontWeight("bold");
          return;
        }

        // Post cleanup
        var dateObj = dt.getValue();
        var shiftTime = shft.getValue();
        var empName = nam.getValue();

        // Set confirmation messages
        const dtFormatted = Utilities.formatDate(dateObj, "GMT+8", "MM/dd");
        const newSheetName = dtFormatted + ' ' + shiftTime + ' ' + empName;
        spreadsheet.getRange(getDupFuncCol() + (endRow + 40) + ':' + getDupFuncCol() + (endRow + 42)).clear({ contentsOnly: true, skipFilteredRows: true });
        spreadsheet.getRange(getDupLabelCol() + (endRow + 43)).setValue('OK').setFontStyle("italic").setFontWeight("bold").setFontColor("green");
        spreadsheet.getRange('B' + (endRow + 45)).setValue('New tab created: "' + newSheetName + '"').setFontStyle("italic").setFontWeight("bold").setFontColor("green");
        actualNewshift(dateObj, shiftTime, empName, propServ, env);

      } else if (a1Not == 'A' + (endRow + 31)) {  // Get delivery
        console.log("Get delivery");
        getDelivery(getStoreCodeByName(sheet.getRange("A1").getValue()), sheet, env);

      } else if (a1Not == 'A' + (endRow + 33)) {  // Hide old -2
        console.log("Hide old sheets");
        hideOldSheets(false, endRow);

      } else if (a1Not == 'A' + (endRow + 35)) {  // Show old
        console.log("Show old sheets");
        var cellLabel = sheet.getRange('A' + (endRow + 34));
        var numSheets = cellLabel.getValue();
        if (isNaN(numSheets)) {
          numSheets = 10;
        }
        cellLabel.setValue("Show old sheets:");
        showLastNSheets(numSheets);

      } else if (a1Not == 'A' + (endRow + 38)) {  // Hide verified
        console.log("Hide unverified sheets");
        hideOldSheets(true, endRow);

      } else if (a1Not == 'A' + (endRow + 40)) {  // Show unverified
        console.log("Show unverified sheets");
        showUnverifiedSheets();

      } else if (a1Not == 'A' + (endRow + 43)) {  // Show unverified
        console.log("Unlock sheet");
        protectDuplicatedSheet(getMBUnprotectedRangeList(), spreadsheet, env);

      } else if (a1Not == 'A' + (endRow + 44)) {  // Auto-formula deliver to ending
        console.log("Auto-formula deliver to ending");
        autoFormulaEnding(spreadsheet);

      } else if (a1Not == getTotalCol() + (endRow + 9)) { // Collect sales button
        console.log("Collecting sales to PO sheet");
        let currentContent = rg.getValue();
        rg.setValue("Processing...");
        rg.offset(0, 1).setValue(Utilities.formatDate(new Date(), "Asia/Hong_Kong", "HH:mm:ss"));
        SpreadsheetApp.flush();

        let totalColIdx = getTotalCol().charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        let totalColVals = sheet.getRange((endRow + 2), totalColIdx, 10, 1).getValues();


        var sheetNameArray = sheet.getSheetName().split(' ');
        var dt = sheetNameArray[0];
        var employeeName = sheetNameArray.slice(2).join(" ");
        // var sales = sheet.getRange(getTotalCol() + (endRow+3)).getValue()          // M36
        // var gcash = sheet.getRange(getTotalCol() + (endRow+4)).getValue()          // M37
        // var expenses = sheet.getRange(getTotalCol() + (endRow+5)).getValue()       // M38
        // var cashAdvance = sheet.getRange(getTotalCol() + (endRow+10)).getValue()   // M43
        // var expectedSales = sheet.getRange(getTotalCol() + (endRow+2)).getValue()  // M35
        // var overLoss = sheet.getRange(getLossOverCol() + (endRow+3)).getValue()
        // var dagdagPeraSaKaha = sheet.getRange(getTotalCol() + (endRow+11)).getValue()
        // var spoiled = sheet.getRange(getTotalCol() + (endRow+6)).getValue()

        let expectedSales = totalColVals[0][0]
        let sales = totalColVals[1][0]
        let gcash = totalColVals[2][0]
        let expenses = totalColVals[3][0]
        let spoiled = totalColVals[4][0]
        let cashAdvance = totalColVals[8][0]
        let dagdagPeraSaKaha = totalColVals[9][0]
        var overLoss = sheet.getRange(getLossOverCol() + (endRow + 3)).getValue();

        var storeName = sheet.getRange("A1").getValue();
        addSalesToCashFlow(storeName, dt, sales, gcash, expenses, cashAdvance, expectedSales, overLoss, employeeName, spoiled, dagdagPeraSaKaha, endRow, getStoreCodeByName(storeName), LockService, env);

        rg.setValue(currentContent);
        SpreadsheetApp.flush();
        rg.check();

      } else if (a1Not == getLossOverCol() + (endRow + 10)) { // Verify delivery
        verifyDelivery(rg, endRow, sheet, env);
      } else if (a1Not == getLossOverCol() + (endRow + 9)) { // Force check the collect checkbox if processing timed-out
        sheet.getRange(getTotalCol() + (endRow + 9)).check();
        sheet.getRange(getLossOverCol() + (endRow + 9)).setValue("Verified");
      } else {
        console.log("Invalid range. endrow=" + endRow + ", rg=" + a1Not);
        let currentContent = rg.getValue();
        rg.setValue("Invalid Checkbox");
        SpreadsheetApp.flush();
        Utilities.sleep(5000);
        rg.setValue(currentContent);
        SpreadsheetApp.flush();
      }

    } else if (a1Not == getDupFuncCol() + (endRow + 43) && (rg.isBlank() || rg.getValue() == "" || !(rg.getValue() == false || rg.getValue() == true))) {   // Prevent checkbox from being deleted
      console.log("Inserting checkboxes");
      rg.insertCheckboxes();

    } else if (a1Not == getDupFuncCol() + (endRow + 41) && (rg.isBlank())) {   // Prevent dropbox from being deleted
      console.log("Setting shift validation");
      rg.setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInList(['AM', 'Mid', 'PM'])
        .build());

    }
  } catch (e) {
    alert(e, "MB RF Inv Err", " [Non-fatal]", env);
    if (!e.stack.includes("already has sheet protection") || !e.stack.includes("failed while accessing document with id") || !e.stack.includes("Timed out")) {
      SpreadsheetApp.getActive().getRange('B' + (endRow + 45)).setValue("ERROR. PLEASE REPORT. \n" + e.stack).setFontColor("red");
    }
    console.error(e.stack);
  }
}
