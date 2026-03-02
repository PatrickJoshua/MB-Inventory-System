function test() {
    let ss = SpreadsheetApp.getActive();
    return;

    let s = ss.getActiveSheet();
    console.log(s.getSheetName() + " " + s.getIndex());

    let lastCol = s.getLastColumn();
    let headerRangeVals = s.getRange(1, 1, 1, lastCol).getValues();
    console.log(headerRangeVals[0][0]);
    //return;

    let sheets = ss.getSheets().filter((sheet) => sheet.getSheetName().startsWith("PO")).slice(350).forEach((sheet) => {
        console.log("Processing " + sheet.getSheetName());
        var dataRange = sheet.getDataRange();
        var lastRow = dataRange.getLastRow();
        var lastColumn = dataRange.getLastColumn();

        // Delete empty rows
        var numRows = sheet.getMaxRows();
        if (numRows > lastRow) {
            sheet.deleteRows(lastRow + 1, numRows - lastRow);
        }

        // Delete empty columns
        var numColumns = sheet.getMaxColumns();
        if (numColumns > lastColumn) {
            sheet.deleteColumns(lastColumn + 1, numColumns - lastColumn);
        }
    });
}

function getGenerateReportCol() {
    return "BE";
}

function getEndRow(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
    return Utils.getRowNum("COLS", spreadsheet);
    //return 41;
}

function scheduledGenerateReport(e) {
    generateReport('', 'M', "3252", 1);
    generateReport('', 'M', "3361", 1);
}

function installedOnEdit(e) {
    const rg = e.range;
    // TODO: CHECK FIRST IF rg,isChecked()
    const spreadsheet = SpreadsheetApp.getActive();
    const a1Not = rg.getA1Notation();   // TODO: FIND AND REPLACE ALL
    const sheetName = spreadsheet.getSheetName();
    try {
        if (rg.getA1Notation() == "U32" && rg.isChecked()) {
            rg.uncheck();
            spreadsheet.getRange('U34').clear();
            nextPO();
            //Utils.triggerFuncWithProcessingText("U34", function() {nextPO()}, spreadsheet)
        } else if (rg.getA1Notation() == "U36" && rg.isChecked()) {
            rg.uncheck();
            Utils.triggerFuncWithProcessingText("U36", function () { pullLatestEnding(spreadsheet.getRange('B6').getValue()); }, spreadsheet);
        } else if (rg.getA1Notation() == "U59" && rg.isChecked()) {
            rg.uncheck();
            sendPO();
        } else if (rg.getA1Notation() == "U75" && rg.isChecked()) {
            rg.uncheck();
            confirmPO();
        } else if ((spreadsheet.getSheetName() == "Report" || spreadsheet.getSheetName() == "Report - PCGH") && rg.getA1Notation() == getGenerateReportCol() + "5" && rg.isChecked()) {  // Generate report
            rg.uncheck();
            try {
                var storeCode = spreadsheet.getRange(getGenerateReportCol() + '1').getValue();
                if (spreadsheet.getRange(getGenerateReportCol() + '4').getValue() == 'Order') {
                    spreadsheet.getRange(getGenerateReportCol() + '7').setFontColor("green").setFontStyle("italic").setFontWeight("bold").setValue("Generating order report...");
                    generateReport('', 'F', storeCode);
                } else if (spreadsheet.getRange(getGenerateReportCol() + '4').getValue() == 'Sales') {
                    spreadsheet.getRange(getGenerateReportCol() + '7').setFontColor("green").setFontStyle("italic").setFontWeight("bold").setValue("Generating sales report...");
                    generateReport('', 'M', storeCode);
                } else {
                    spreadsheet.getRange(getGenerateReportCol() + '7').setFontColor("red").setFontStyle("italic").setFontWeight("bold").setValue("Please select report type");
                    Utilities.sleep(3000);
                }
                spreadsheet.getRange(getGenerateReportCol() + '7').clear();
            } catch (e) {
                spreadsheet.getRange(getGenerateReportCol() + '7').setFontColor("red").setFontStyle("italic").setFontWeight("bold").setValue(e.stack);
            }
        } else if (rg.getA1Notation() == "O77" && rg.isChecked()) {
            rg.uncheck();
            spreadsheet.getRange('P77').setValue('');
            addPoToCashFlow(spreadsheet.getRange("B6").getValue());
        } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "Q1" && rg.isChecked()) {
            rg.uncheck();
            computeTotalCashCollected(spreadsheet);
        } else if ((rg.getA1Notation() == "R65" || rg.getA1Notation() == "V74") && rg.isChecked()) {
            rg.uncheck();
            Utils.incrementLeftCell(spreadsheet, rg.getA1Notation());
        } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "S3" && rg.isChecked()) {
            rg.uncheck();
            appendToCashReceived(spreadsheet.getRange("R3").getValue(), "R3", spreadsheet);
        } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "S4" && rg.isChecked()) {
            rg.uncheck();
            appendToExpenses(5, spreadsheet);
        } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "R1" && rg.isChecked()) {
            rg.uncheck();
            appendToCashCollected(spreadsheet);
        } else if (spreadsheet.getSheetName().startsWith("RF/PCGH") && rg.getA1Notation() == "A1" && rg.isChecked()) {
            rg.uncheck();
            Utils.triggerFuncWithProcessingText("D1", function () { pattyDistribution(spreadsheet); }, spreadsheet);
        } else if (spreadsheet.getSheetName() == "InventoryReplica" && rg.getA1Notation() == "A1" && rg.isChecked()) {
            rg.uncheck();
            Utils.triggerFuncWithProcessingText("C1", function () { updateInventoryReplica(spreadsheet.getActiveSheet()); }, spreadsheet);
        } else if (spreadsheet.getSheetName() == "GCash" && (rg.getA1Notation() == "C1" || rg.getA1Notation() == "I1") && rg.isChecked()) {
            rg.uncheck();
            let sheet = spreadsheet.getActiveSheet();
            //let labelRg = sheet.getRange(rg.getRow(), rg.getColumn()+1)
            //Utils.triggerFuncWithProcessingText(labelRg.getA1Notation(), function() {addGcashToCashReceived(rg, sheet)}, sheet)
            Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { addGcashToCashReceived(rg, sheet); }, sheet);
        } else if (spreadsheet.getSheetName() == "GCash" && rg.getA1Notation() == "L1" && rg.isChecked()) { // Trigger collect gcash on both stores
            rg.uncheck();
            let sheet = spreadsheet.getActiveSheet();
            Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { addGcashToCashReceived(sheet.getRange("C1"), sheet); }, sheet);
            Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { addGcashToCashReceived(sheet.getRange("I1"), sheet); }, sheet);
        } else if (spreadsheet.getSheetName() == "GCash" && (rg.getA1Notation() == "E1" || rg.getA1Notation() == "K1") && rg.isChecked()) {
            rg.uncheck();
            let sheet = spreadsheet.getActiveSheet();
            //let labelRg = sheet.getRange(rg.getRow(), rg.getColumn()+1)
            Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { updateInventoryReplica(spreadsheet.getSheetByName("InventoryReplica")); }, sheet);
        } else if (rg.isChecked() && sheetName.startsWith("*")) {
            //Utils.triggerFuncWithProcessingText(a1Not, function() { proxyAddToCashflow(e, a1Not, spreadsheet) }, spreadsheet)
            proxyAddToCashflow(e, a1Not, spreadsheet);
        } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "S1" && rg.isChecked()) {
            rg.uncheck();
            Utils.triggerFuncWithProcessingText(a1Not, getUnverifiedSheets, spreadsheet);
        }
    } catch (e) {
        spreadsheet.getRange('U34').setValue(e).setFontColor("red");
        console.log(e.stack);
        Utils.alert(e, "MB PO Err");
        throw e;
    }
}



function pullLatestEnding(storeCode, sheet = SpreadsheetApp.getActive().getActiveSheet(), isGenerateReport = true) {
    console.log("Pulling latest ending for store " + storeCode);
    var inventorySpreadsheet = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode));
    var inventorySheets = inventorySpreadsheet.getSheets();
    var latestInvSheet = inventorySheets[inventorySheets.length - 1];
    const endRow = Utils.getEndRow(latestInvSheet);

    var stocksCol = "D";
    var currentStocksCell = Utils.getLastCol() + (endRow + 34);
    var currentSheetStocksValue = latestInvSheet.getRange(currentStocksCell).getValue();
    console.log("Current sheet's stocks value: " + currentSheetStocksValue + " from " + currentStocksCell);
    var generateReportOffset = 0;
    if (currentSheetStocksValue == 0) {
        console.log("No ending detected. Using beginning stocks.");
        stocksCol = "B";
        generateReportOffset = 1;
    }

    [[stocksCol, "Z"], ["A", "X"], ["J", "Y"]].forEach((pair) => {
        var invCol = pair[0];
        var poCol = pair[1];
        sheet.getRange(poCol + '2:' + poCol + endRow).setValues(latestInvSheet.getRange(invCol + '2:' + invCol + endRow).getValues());
    });
    // var ending = latestInvSheet.getRange('D2:D' + endRow).getValues();
    // console.log(ending.toString())
    // spreadsheet.getRange('Z2:Z' + endRow).setValues(ending)
    if (isGenerateReport)
        generateReport('', 'M', storeCode, generateReportOffset);

    return latestInvSheet;
}

function nextPO(storeCode) {
    console.log("Generating next PO for store " + storeCode);
    var spreadsheet = SpreadsheetApp.getActive();

    // Check first if PO is registered to cashflow
    if (spreadsheet.getRange('P77').isBlank()) {
        spreadsheet.getRange("U34").setFontColor("red").setFontWeight("bold").setFontStyle("italic").setValue("ERROR: Current PO is not yet added to cash flow");
        SpreadsheetApp.flush();
        Utilities.sleep(5000);
        spreadsheet.getRange("U34").clearContent();
        return;
    }

    var nextPoDate = new Date(spreadsheet.getRange('O35').getValue());
    var days2consume = spreadsheet.getRange('O37').getValue();
    nextPoDate.setDate(nextPoDate.getDate() + days2consume);

    var storeCode = spreadsheet.getRange("B6").getValue();

    const dtFormatted = Utilities.formatDate(nextPoDate, "GMT+8", "MM/dd/yy");
    const prevSheet = spreadsheet.getActiveSheet();
    spreadsheet.duplicateActiveSheet();
    spreadsheet.getActiveSheet().setName('PO D' + dtFormatted + " " + storeCode);

    spreadsheet.getRange('O35').setValue(nextPoDate);

    // Copy to previous order
    spreadsheet.getRange('I31:I').copyTo(spreadsheet.getRange('J31'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    // Cleanup
    spreadsheet.getRange('I31:I').setValue(''); // Prev Orders
    spreadsheet.getRange('O66:O72').setValue(''); // PO/SO Confirmation
    spreadsheet.getRange('P77').setValue(''); // Cash flow added text
    spreadsheet.getRange('P82').setValue(''); // Prev actual PO amount
    spreadsheet.getRange('U63').setValue(''); // Returned SMS
    spreadsheet.getRange('U74').setValue(''); // Returned SMS position
    spreadsheet.getRange('Q65').setValue(''); // Confirmation SMS position

    // Hardcode previous projected sales
    const prevSheetProjSalesRg = prevSheet.getRange("G7");
    prevSheetProjSalesRg.setValue(prevSheetProjSalesRg.getValue());

    pullLatestEnding(storeCode, spreadsheet, true);
    //generateReport('', 'M', storeCode);
    prevSheet.hideSheet();
}

function generateReport(reportName, column, storeCode, rightOffset = 0) {
    console.log("Generating report for store " + storeCode);
    var inventorySpreadsheet = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode));
    var inventorySheets = inventorySpreadsheet.getSheets();
    //var inventorySheets = _inventorySheets//.slice(60)

    // Store Switch
    var reportNameFinal = reportName + "Report";
    if (storeCode == "3361") {
        reportNameFinal = reportNameFinal + " - PCGH";
    }
    console.log("Report name: " + reportNameFinal);
    var reportSheet = SpreadsheetApp.getActive().getSheetByName(reportNameFinal);
    // var rowIt = 2

    var isColumnNamesEstablished = false;
    //console.log(reportSheet.getRange("A:A").getValues().filter(String))
    var lastRow = reportSheet.getRange("A:A").getValues().filter(String).length + 1;
    if (!reportSheet.getRange("A1").isBlank()) {
        isColumnNamesEstablished = true;
    } else {
        lastRow++;
    }

    console.log("Last row: " + lastRow);
    var inventorySheetsLength = inventorySheets.length;
    console.log("Number of inventories: " + inventorySheetsLength);

    // Archive old sheets
    const sheetsToRetain = 30; //inventorySheetsLength-20 //30
    var archiveSpreadsheet = SpreadsheetApp.openByUrl(Utils.getArchiveInventoryUrl(storeCode));
    if (inventorySheetsLength > sheetsToRetain + 2) {
        let today = new Date();
        let year = today.getFullYear() - 2000;
        let isCurrentMonthJan = today.getMonth() == 0 ? true : false;
        for (i = 2; i < inventorySheetsLength - sheetsToRetain; i++) {
            const currentSheetName = inventorySheets[i].getSheetName();
            let nameArray = currentSheetName.split(" ");

            if (isCurrentMonthJan && nameArray[0].startsWith("12/")) { // check if crossing New year with remaining December inventories to be archived
                nameArray[0] = nameArray[0] + "/" + (year - 1);
            } else {
                nameArray[0] = nameArray[0] + "/" + year;
            }
            let newSheetName = nameArray.join(" ");

            console.log("Archiving sheet: " + newSheetName);
            // Pre-check if existing
            let existingSheet = archiveSpreadsheet.getSheetByName(newSheetName);
            if (existingSheet) {
                console.log(`Sheet ${newSheetName} already exists in archive. Skipping.`);
            } else {
                let copiedArchiveSheet;
                try {
                    copiedArchiveSheet = inventorySheets[i].copyTo(archiveSpreadsheet);
                    copiedArchiveSheet.setName(newSheetName);
                } catch (e) {
                    console.log(e.stack);

                    if (e.stack.includes("already exists. Please enter another name.")) {
                        archiveSpreadsheet.deleteSheet(copiedArchiveSheet);
                    } else {
                        throw e;    // Temporarily halt execution until handling design has been finalized
                    }
                }
            }

            inventorySpreadsheet.deleteSheet(inventorySheets[i]);
        }
    }

    var inventorySheetsLength = inventorySpreadsheet.getNumSheets();
    var archiveSheetsLength = archiveSpreadsheet.getNumSheets() - 2;
    var inventorySheets = inventorySpreadsheet.getSheets();
    var archiveSheets = archiveSpreadsheet.getSheets();
    console.log("inventory+archive " + inventorySheetsLength + "+" + archiveSheetsLength + " = " + (inventorySheetsLength + archiveSheetsLength));
    console.log("lastRow<=inventory+archive-1 " + lastRow + " <= " + (inventorySheetsLength + archiveSheetsLength - 1));

    // Force update latest
    if (lastRow - 1 == inventorySheetsLength + archiveSheetsLength - 1) {
        console.log("Force updating row " + --lastRow);
        //sheet.deleteRow(--lastRow);
    }

    // inventorySheets.forEach(function(sheet) {
    for (j = lastRow; j <= (inventorySheetsLength + archiveSheetsLength - 1 - rightOffset); j++) {
        let sheet = j - archiveSheetsLength < 2 ? archiveSheets[j] : inventorySheets[j - archiveSheetsLength];  // Traverse archiveSheets if already archived
        console.log("current iterator:" + j + "; current index: " + (j - archiveSheetsLength));
        const endRow = getEndRow(sheet);
        let sheetName = sheet.getName();
        let splittedName = sheetName.split(' ');
        // if (splittedName[0].search('^[0-9]{1,2}\/[0-9]{1,2}$') != -1) {
        // Setup column names
        if (!isColumnNamesEstablished) {
            reportSheet.getRange(1, 1, 1, 3).setValues([['Date', 'Shift', 'Employee']]);

            let products = inventorySheets[inventorySheets.length - 1].getRange('A2:A' + endRow).getValues();
            let transposedProducts = products[0].map((_, colIndex) => products.map((row) => row[colIndex]));
            console.log("Column name setup: " + transposedProducts);
            reportSheet.getRange(1, 4, 1, endRow - 1).setValues(transposedProducts);

            isColumnNamesEstablished = true;
        }

        console.log("Extracting: " + sheetName);
        let dt = splittedName[0];
        let time = splittedName[1];
        let name = splittedName.slice(2).join(' ');

        reportSheet.getRange(j, 1, 1, 3).setValues([[dt, time, name]]);

        // var _column = column;


        // if (column == "M" && sheet.getRange("K1").getValue() == "Sales") { // legacy Sales column position
        //   _column = "K"
        // }
        orders = sheet.getRange(column + '2:' + column + endRow).getValues();
        transposedOrders = orders[0].map((_, colIndex) => orders.map((row) => row[colIndex]));
        reportSheet.getRange(j, 4, 1, endRow - 1).setValues(transposedOrders);

        // rowIt++;
        //}
    }

    // Hide outdated products
    //reportSheet.hideColumns(12,2)
    //reportSheet.hideColumns(22)
}

function addPoToCashFlow(storeCode) {
    var spreadsheet = SpreadsheetApp.getActive();
    var confirmationText = spreadsheet.getRange('O66').getValue();
    const dt = Utilities.formatDate(new Date(spreadsheet.getRange('O35').getValue()), "GMT+8", "MM/dd");

    try {
        var amt = confirmationText.match('(AMT=)(.*)( is)')[2];
        var so = confirmationText.match('(SO=)(.*)( AMT)')[2];
        var captureStoreCodeMatch = confirmationText.match('(Your order for )(\\\\d\\\\d\\\\d\\\\d)(.*)( w/ SO=)');
        var capturedStoreCode = captureStoreCodeMatch[2];
    } catch (e) {
        spreadsheet.getRange('P77').setFontColor("red").setFontWeight("bold").setFontStyle("italic").setValue("Confirmation text is not valid.\n" + e.stack);
        console.error(e.stack);
        return;
    }

    // Check if correct store PO
    if (storeCode != capturedStoreCode) {
        spreadsheet.getRange('P77').setFontColor("red").setFontWeight("bold").setFontStyle("italic").setValue("PO's store code is not for this sheet. Expected: " + storeCode + "; Got: " + capturedStoreCode);
        return;
    }

    // Set value to actual PO cell
    spreadsheet.getRange('P82').setValue(amt);

    // Store Switch
    var cashFlowSheetName = "Cash flow";
    if (storeCode == "3361") {
        cashFlowSheetName = "Cash flow - PCGH";
    }

    // Append value
    var cashFlowSheet = spreadsheet.getSheetByName(cashFlowSheetName);
    var colValues = cashFlowSheet.getRange("A:A").getValues();
    var count = colValues.filter(String).length;
    cashFlowSheet.getRange(count + 1, 1).setValue(amt);
    cashFlowSheet.getRange(count + 1, 2).setValue(dt + ' - ' + so);

    spreadsheet.getRange('P77').setFontColor("green").setFontWeight("bold").setFontStyle("italic").setValue("Added " + amt + " to Cash flow sheet");
}

function sendPO(smsApiUrl = "https://docs.google.com/spreadsheets/d/17yPemlid9FVMdzVDX8Eg8Tu1W-zOg_prNtQeUeEidAg/edit") {
    console.log("Sending PO");
    var spreadsheet = SpreadsheetApp.getActive();
    var poNum = spreadsheet.getRange("U58").getValue();
    var poStr = spreadsheet.getRange("O58").getValue();
    var smsApiSheet = SpreadsheetApp.openByUrl(smsApiUrl).getSheetByName("SMS");
    smsApiSheet.appendRow([poNum, poStr, true]);

    var lastRow = smsApiSheet.getLastRow() + 1;
    const importRange = 'IMPORTRANGE("' + smsApiUrl + '", "' + "'SMS'!B\"&U74" + ')';
    console.log(importRange + "\nSMS Row: " + lastRow);
    spreadsheet.getRange("U74").setValue(lastRow);
    spreadsheet.getRange("U63").setFormula(importRange);
}

function confirmPO(smsApiUrl = "https://docs.google.com/spreadsheets/d/17yPemlid9FVMdzVDX8Eg8Tu1W-zOg_prNtQeUeEidAg/edit") {
    console.log("Confirming PO");
    var spreadsheet = SpreadsheetApp.getActive();
    var smsRow = spreadsheet.getRange("U74").getValue();
    var smsApiSheet = SpreadsheetApp.openByUrl(smsApiUrl).getSheetByName("SMS");
    smsApiSheet.getRange("'SMS'!C" + smsRow).setFormula(true);

    var lastRow = smsApiSheet.getLastRow() + 1;
    const importRange = 'IMPORTRANGE("' + smsApiUrl + '", "' + "'SMS'!B\"&Q65" + ')';
    console.log(importRange + "\nSMS Row: " + lastRow);
    spreadsheet.getRange("Q65").setValue(lastRow);
    spreadsheet.getRange("O66").setFormula(importRange);
}

function computeTotalCashCollected(sheet = SpreadsheetApp.getActiveSheet()) {
    var colValues = sheet.getRange("H:H").getValues();
    var cashLastRow = colValues.filter(String).length + 1;
    var collectLastRow = sheet.getLastRow();
    var range = sheet.getRange("O" + collectLastRow);
    if (range.isBlank() || range.getValue() == "") {
        collectLastRow = range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    }

    // Get range of previous cash advance
    //let previousLastRow = sheet.getRange("O:O").getValues().filter(String).length

    sheet.getRange("O" + cashLastRow).setFormula("SUM(H" + (collectLastRow + 1) + ":H" + cashLastRow + ")-Q" + collectLastRow + "+Q" + cashLastRow);
    sheet.getRange("P" + cashLastRow).setFormula("SUM(M" + (collectLastRow + 1) + ":M" + cashLastRow + ")");
    //sheet.getRange("Q" + cashLastRow).setFormula("(1000*0)+(500*0)+(200*0)+(100*0)+(50*0)+(20*0)+(10*0)+(5*0)+0-O" + cashLastRow)
}

function appendToCashReceived(amt, rg = null, sheet = SpreadsheetApp.getActiveSheet()) {
    if (!amt) return;
    var colValues = sheet.getRange("F:F").getValues();
    var count = colValues.filter(String).length;
    sheet.getRange("E" + (count + 2)).setValue(new Date());
    sheet.getRange("F" + (count + 2)).setValue(amt);

    // Print
    if (rg != null) {
        sheet.getRange(rg).setValue("Added " + amt + " to cash received");
        SpreadsheetApp.flush();
        Utilities.sleep(5000);
        sheet.getRange(rg).setValue("");
    }

}

function appendToExpenses(startingRow, sheet = SpreadsheetApp.getActiveSheet()) {
    var rawExpenseSheetName = sheet.getSheetName().replace("Cash flow", "Raw Expenses");
    var rawExpenseSheet = SpreadsheetApp.getActive().getSheetByName(rawExpenseSheetName);

    let validExpenses = [];
    let cashReceiveds = [];
    let rawExpenses = [];
    for (i = startingRow; i < 12; i++) {
        console.log("Reading R" + i);
        var expenseAmount = sheet.getRange("R" + i).getValue();
        console.log("Expense name: " + expenseName);
        if (expenseAmount) {
            console.log("Valid expense");
            // Get Expenses last row
            var expenseName = sheet.getRange("Q" + i).getValue();
            var expValues = sheet.getRange("H:H").getValues();
            var expCount = expValues.filter(String).length;

            // Get Cash received last row
            var receivedValues = sheet.getRange("F:F").getValues();
            var receivedCount = receivedValues.filter(String).length;

            // Append to Expenses
            sheet.getRange("G" + (expCount + 2)).setValue(new Date());
            sheet.getRange("H" + (expCount + 2)).setValue(0);
            sheet.getRange("J" + (expCount + 2)).setValue(expenseAmount);
            sheet.getRange("N" + (expCount + 2)).setValue(expenseName);
            validExpenses.push([new Date(), 0, expenseAmount, expenseName]);

            // Append to Cash received
            sheet.getRange("E" + (receivedCount + 2)).setValue(new Date());
            sheet.getRange("F" + (receivedCount + 2)).setFormula("-J" + (expCount + 2));
            cashReceiveds.push([new Date(), "-J" + (expCount + 2)]);

            // Append to Raw Expenses
            console.log("[DEBUG] rawExpenseSheet=" + rawExpenseSheet.getSheetName());
            var colValues = rawExpenseSheet.getRange("A:A").getValues();
            var count = colValues.filter(String).length;
            console.log("rawExpenseSheet last row: " + count);
            // rawExpenseSheet.getRange("A" + (count+1)).setValue(new Date());
            // rawExpenseSheet.getRange("C" + (count+1)).setValue(expenseName);
            // rawExpenseSheet.getRange("D" + (count+1)).setValue(expenseAmount);
            rawExpenseSheet.getRange((count + 1), 1, 1, 4).setValues([[new Date(), null, expenseName, expenseAmount]]);
            rawExpenses.push([new Date(), null, expenseName, expenseAmount]);
            rawExpenseSheet.getRange("E" + (count)).copyTo(rawExpenseSheet.getRange("E" + (count + 1)));

            //sheet.getRange("R" + i).setValue("Expensed " + sheet.getRange("F" + (receivedCount+2)).getValue())
            //SpreadsheetApp.flush();
        }
    }

    for (i = 0; i < validExpenses.length; i++) {
        break;
    }

    SpreadsheetApp.flush();
    Utilities.sleep(5000);
    sheet.getRange("R5:R11").setValue("");
    sheet.getRange("Q5").setValue("");
}

function appendToCashCollected(sheet = SpreadsheetApp.getActiveSheet()) {
    var colValues = sheet.getRange("H:H").getValues();
    var count = colValues.filter(String).length;

    var receivedValues = sheet.getRange("F:F").getValues();
    var receivedCount = receivedValues.filter(String).length;

    sheet.getRange("E" + (receivedCount + 2)).setValue(new Date());
    sheet.getRange("F" + (receivedCount + 2)).setFormula("R" + (count + 1));

    // Print
    //var addedVal = sheet.getRange("F" + (receivedCount+2)).getValue();
    let addedVal = sheet.getRange("O" + (count + 1)).getValue();
    sheet.getRange("R" + (count + 1)).setValue(addedVal);
    sheet.getRange("S" + (count + 1)).setValue("Added to cash received");
    SpreadsheetApp.flush();
}

function extractExpensesLoop(storeCode = "3361") {
    var inventorySpreadSheet = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode));
    console.log("Spreadsheet name: " + inventorySpreadSheet.getName());
    var sheets = inventorySpreadSheet.getSheets();
    var props = PropertiesService.getScriptProperties();
    x = parseInt(props.getProperty("expenseIndexCounter"));
    for (i = x; i < x + 200; i++) {
        var inventorySheet = sheets[i];
        var inventorySheetName = inventorySheet.getSheetName();
        var split = inventorySheetName.split(" ");
        var dt = split[0];
        var employeeName = split[2];
        //console.log("Extracting expenses on: " + inventorySheetName)
        Utils.extractExpenses(dt, employeeName, storeCode, inventorySheet);
        //console.log("current index: " + i)
        props.setProperty("expenseIndexCounter", i);
        SpreadsheetApp.flush();
    }
}

function adhocExtractExpensesFromCashFlow() {
    var sheet = SpreadsheetApp.getActive().getSheetByName("Cash flow - PCGH");
    var rawExpenseSheetName = sheet.getSheetName().replace("Cash flow", "Raw Expenses");
    var rawExpenseSheet = SpreadsheetApp.getActive().getSheetByName(rawExpenseSheetName);
    var colValues = rawExpenseSheet.getRange("A:A").getValues();
    var count = colValues.filter(String).length;

    for (i = 2; i < 393; i++) {
        var sales = sheet.getRange("K" + i).getValue();
        if (!sales) {
            rawExpenseSheet.getRange("A" + (++count)).setValue(sheet.getRange("G" + i).getValue());
            rawExpenseSheet.getRange("C" + (count)).setValue(sheet.getRange("N" + i).getValue());
            rawExpenseSheet.getRange("D" + (count)).setValue(sheet.getRange("J" + i).getValue());
        }
    }
}

function pattyDistribution(spreadsheet = SpreadsheetApp.getActive()) {
    console.log("Update Patty distribution");

    Utils.getStoreCodes().forEach((storeCode) => {
        console.log("Current store code: " + storeCode);

        var lastPoSheet = Utils.getLastPoSheet(storeCode);

        // pull latest
        // pullLatestEnding(storeCode, lastPoSheet, false)

        // populate ordered sets
        for (i = 1; i < spreadsheet.getLastRow(); i++) {
            //console.log("[DEBUG] looking on row " + i + ": " + spreadsheet.getRange("A" + i).getValue())
            if (spreadsheet.getRange("A" + i).getValue() == storeCode) {
                console.log("Found on row: " + i);
                //spreadsheet.getRange("E" + i).setValue(lastPoSheet.getspreadsheetName()) // Set spreadsheet name for stocks lookup

                var poMap = Utils.constructPoMap(lastPoSheet);
                while (true) {
                    var product = spreadsheet.getRange("A" + ++i).getValue();
                    console.log("Product: " + product);
                    if (!product) break;
                    spreadsheet.getRange("C" + i).setValue(poMap.get(product));
                }

                // Clear previous patty distribution
                while (spreadsheet.getRange("B" + (i)).getValue() != "Freezer Top") { console.log("Looking for Freezer Top on row " + i++); }
                while (spreadsheet.getRange("B" + ++i).getValue() != "Freezer Bottom") {
                    console.log("Clearing row " + i);
                    spreadsheet.getRange("B" + i).setValue("");
                    spreadsheet.getRange("E" + i).setValue("");
                }
                break;
            }
        }
    });

    // Update Inventory Replica
    updateInventoryReplica(spreadsheet.getSheetByName("InventoryReplica"));
}

function updateInventoryReplica(sheet = SpreadsheetApp.getActive().getActiveSheet()) {
    // let map = new Map();
    // map.set("3252", "A2")
    // map.set("3361", "P2")

    let storeCodes = Utils.getStoreCodes();
    let processedStoreCodes = 0;

    for (i = 1; i < sheet.getLastColumn(); i++) {
        let storeCode = sheet.getRange(2, i).getValue();
        //console.log("Scanning " + storeCode)
        if (storeCodes.includes(String(storeCode))) {
            console.log("Found " + storeCode);
            let sheets = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode)).getSheets();
            let lastSheetName = sheets[sheets.length - 1].getSheetName();
            sheet.getRange(2, ++i).setValue(lastSheetName);
            processedStoreCodes++;
        }

        if (processedStoreCodes >= storeCodes.length) break;
    }

    // Utils.getStoreCodes().forEach((storeCode) => {
    //   let sheets = sheetApp.openByUrl(Utils.getInventoryUrl(storeCode)).getSheets();
    //   let lastSheetName = sheets[sheets.length-1].getSheetName();
    //   sheet.getRange(map.get(storeCode)).setValue(lastSheetName);
    // })
}

function addGcashToCashReceived(rg, gcashSheet = SpreadsheetApp.getActive().getActiveSheet()) {
    let rgRow = rg.getRow();
    let rgCol = rg.getColumn();

    let startRow = rgRow + 2;
    let shiftCol = rgCol - 2;
    let gcashCol = rgCol - 1;
    let manualCol = rgCol + 1;
    let replicaCol = rgCol + 2;

    let totalGcashRg = gcashSheet.getRange(rgRow, rgCol - 1);
    let totalGcashVal = totalGcashRg.getValue();
    let actualGcashRg = gcashSheet.getRange(rgRow, rgCol + 1);
    let actualGcashVal = totalGcashVal;

    // Replace the var to be recorded to the actual gcash amount if supplied
    if (!actualGcashRg.isBlank()) {
        actualGcashVal = actualGcashRg.getValue();
    }

    // Initialize store name
    let storeName = gcashSheet.getRange(rgRow, shiftCol).getValue();
    console.log("Store name: " + storeName);

    // should yield 0 if no over/loss
    let gcashVariance = actualGcashVal - totalGcashVal;
    console.log(`GCash variance: ${gcashVariance}`);
    if (gcashVariance != 0) {
        Utils.cashCollectedAppender(storeName, `${new Date().getMonth() + 1}/${new Date().getDate()}/${new Date().getFullYear()}`, 0, 0, 0, 0, 0, 0, gcashVariance, "GCash", 0);
    }

    // Save actual gcash amount to cash received
    appendToCashReceived(actualGcashVal, null, Utils.getCashFlowSheet(storeName));

    // Clear
    gcashSheet.getRange(startRow, shiftCol, 998, 2).clear();
    gcashSheet.getRange(startRow, rgCol, 998, 1).uncheck();
    actualGcashRg.clear(); // clear actual gcash range

    // Save state of start row for manual after replica mutation
    let manualStartRow = startRow;
    let replicaAndManualCounter = 0;

    // Copy replicas
    while (true) {
        let manualRg = gcashSheet.getRange(startRow, replicaCol);
        let manualVal = manualRg.getValue();

        if (manualVal) {
            gcashSheet.getRange(startRow, gcashCol).setValue(manualVal);
            gcashSheet.getRange(startRow, rgCol).check();
            startRow++;
            replicaAndManualCounter++;
        } else {
            break;
        }
    }

    // Move manuals
    while (true) {
        let manualRg = gcashSheet.getRange(manualStartRow, manualCol);
        let manualVal = manualRg.getValue();

        if (manualVal) {
            gcashSheet.getRange(startRow, gcashCol).setValue(manualVal);
            gcashSheet.getRange(startRow, rgCol).check();
            manualRg.clear();
            startRow++;
            manualStartRow++;
            replicaAndManualCounter++;
        } else {
            break;
        }
    }

    // Auto-check future replica and manual entries
    while (replicaAndManualCounter-- > 0) {
        gcashSheet.getRange(startRow++, rgCol).check();
    }

    // Note: optimize the loops above to minimize getRange(), getValue(), clear(), and check() operations
}

function getUnverifiedSheets(sheet = SpreadsheetApp.getActiveSheet()) {
    let storeCode = sheet.getRange("A1").getValue();
    let inventorySheet = Utils.getPoSpreadsheet(Utils.getInventoryUrl(storeCode));
    let unverifiedSheets = Utils.showUnverifiedSheets(inventorySheet);
    let spreadsheet = sheet.getParent();

    unverifiedSheets.forEach((unverifiedSheet) => {
        let unverifiedSheetName = unverifiedSheet.getSheetName();
        console.log("Copying sheet to PO spreadsheet: " + unverifiedSheetName);

        let newSheetName = "*" + unverifiedSheetName;
        unverifiedSheet.copyTo(spreadsheet).setName(newSheetName);

        let newSheet = spreadsheet.getSheetByName(newSheetName);
        let rangeStr = "A1:N" + unverifiedSheet.getLastRow();
        newSheet.getRange(rangeStr).setValues(unverifiedSheet.getRange(rangeStr).getValues());   // Flatten formulas
    });
}

function proxyAddToCashflow(e, a1Not, spreadsheet = SpreadsheetApp.getActive()) {
    let sheet = spreadsheet.getActiveSheet();
    let endRow = Utils.getEndRow(sheet);

    if (a1Not == Utils.getTotalCol() + (endRow + 9)) {
        //sheet.getRange(a1Not).setValue("Processing...")
        sheet.setName(sheet.getSheetName().substring(1));  // Remove the asterisk

        let inventorySpreadsheet = Utils.getPoSpreadsheet(Utils.getInventoryUrl(Utils.getStoreCodeByName(sheet.getRange("A1").getValue())));
        let inventorySheet = inventorySpreadsheet.getSheetByName(sheet.getSheetName());
        console.log(`Referenced inventory name: ${inventorySheet.getSheetName()}`);

        //let inventoryRg = inventorySheet.getRange(a1Not);
        //inventoryRg.check();   // Check the box for verify in the original inventory sheet
        inventorySheet.getRange(a1Not).check();
        let modifiedEvent = { e, range: inventorySheet.getRange(a1Not), source: inventorySpreadsheet };
        //console.log(e.range.isChecked());

        Utils.installedOnEditTrigger(modifiedEvent); // Trigger the inventory function
        spreadsheet.deleteSheet(sheet);  // Delete the replica sheet on PO
    }
}

function autoUpdateInventoryReplica(e) {
    updateInventoryReplica(SpreadsheetApp.getActive().getSheetByName("InventoryReplica"));
}
