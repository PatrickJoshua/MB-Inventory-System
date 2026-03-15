const PO = {
    testPO: () => {
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
    },

    getGenerateReportCol: () => {
        return "BE";
    },

    getEndRow: (spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) => {
        return Utils.getRowNum("COLS", spreadsheet);
        //return 41;
    },

    scheduledGenerateReport: (e) => {
        PO.generateReport('', 'M', "3252", 1);
        PO.generateReport('', 'M', "3361", 1);
    },

    installedOnEditTriggerPO: (e, propServ = PropertiesService, env = 'PRD') => {
        // Priority Check: Use env from PropertiesService if available
        const savedEnv = propServ.getScriptProperties().getProperty("env");
        if (savedEnv) {
            console.log(`[INFO] Environment override: Using "${savedEnv}" from PropertiesService (Argument was "${env}")`);
            env = savedEnv;
        } else {
            console.log(`[INFO] Using environment: "${env}" (No override found in PropertiesService)`);
        }

        const rg = e.range;
        // TODO: CHECK FIRST IF rg,isChecked()
        const spreadsheet = SpreadsheetApp.getActive();
        const a1Not = rg.getA1Notation();   // TODO: FIND AND REPLACE ALL
        const sheetName = spreadsheet.getSheetName();
        try {
            if (rg.getA1Notation() == "U32" && rg.isChecked()) {
                rg.uncheck();
                spreadsheet.getRange('U34').clear();
                PO.nextPO(Utils.getStoreCodeByName(spreadsheet.getRange("B6").getValue()), env);
                //Utils.triggerFuncWithProcessingText("U34", function() {PO.nextPO()}, spreadsheet)
            } else if (rg.getA1Notation() == "U36" && rg.isChecked()) {
                rg.uncheck();
                Utils.triggerFuncWithProcessingText("U36", function () { PO.pullLatestEnding(spreadsheet.getRange('B6').getValue(), undefined, true, env) }, spreadsheet);
            } else if (rg.getA1Notation() == "U59" && rg.isChecked()) {
                rg.uncheck();
                PO.sendPO(env);
            } else if (rg.getA1Notation() == "U75" && rg.isChecked()) {
                rg.uncheck();
                PO.confirmPO(env);
            } else if ((spreadsheet.getSheetName() == "Report" || spreadsheet.getSheetName() == "Report - PCGH") && rg.getA1Notation() == PO.getGenerateReportCol() + "5" && rg.isChecked()) {  // Generate report
                rg.uncheck();
                try {
                    var storeCode = spreadsheet.getRange(PO.getGenerateReportCol() + '1').getValue();
                    if (spreadsheet.getRange(PO.getGenerateReportCol() + '4').getValue() == 'Order') {
                        spreadsheet.getRange(PO.getGenerateReportCol() + '7').setFontColor("green").setFontStyle("italic").setFontWeight("bold").setValue("Generating order report...");
                        PO.generateReport('', 'F', storeCode, 0, env);
                    } else if (spreadsheet.getRange(PO.getGenerateReportCol() + '4').getValue() == 'Sales') {
                        spreadsheet.getRange(PO.getGenerateReportCol() + '7').setFontColor("green").setFontStyle("italic").setFontWeight("bold").setValue("Generating sales report...");
                        PO.generateReport('', 'M', storeCode, 0, env);
                    } else {
                        spreadsheet.getRange(PO.getGenerateReportCol() + '7').setFontColor("red").setFontStyle("italic").setFontWeight("bold").setValue("Please select report type");
                        Utilities.sleep(3000);
                    }
                    spreadsheet.getRange(PO.getGenerateReportCol() + '7').clear();
                } catch (e) {
                    spreadsheet.getRange(PO.getGenerateReportCol() + '7').setFontColor("red").setFontStyle("italic").setFontWeight("bold").setValue(e.stack);
                }
            } else if (rg.getA1Notation() == "O77" && rg.isChecked()) {
                rg.uncheck();
                spreadsheet.getRange('P77').setValue('');
                PO.addPoToCashFlow(spreadsheet.getRange("B6").getValue(), env);
            } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "Q1" && rg.isChecked()) {
                rg.uncheck();
                PO.computeTotalCashCollected(spreadsheet);
            } else if ((rg.getA1Notation() == "R65" || rg.getA1Notation() == "V74") && rg.isChecked()) {
                rg.uncheck();
                Utils.incrementLeftCell(spreadsheet, rg.getA1Notation());
            } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "S3" && rg.isChecked()) {
                rg.uncheck();
                PO.appendToCashReceived(spreadsheet.getRange("R3").getValue(), "R3", spreadsheet);
            } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "S4" && rg.isChecked()) {
                rg.uncheck();
                PO.extractExpensesLoop(Utils.getStoreCodeByName(spreadsheet.getRange("A1").getValue()), env);
            } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "R1" && rg.isChecked()) {
                rg.uncheck();
                PO.appendToCashCollected(spreadsheet);
            } else if (spreadsheet.getSheetName().startsWith("RF/PCGH") && rg.getA1Notation() == "A1" && rg.isChecked()) {
                rg.uncheck();
            } else if (spreadsheet.getSheetName() == "InventoryReplica" && rg.getA1Notation() == "A1" && rg.isChecked()) {
                rg.uncheck();
                Utils.triggerFuncWithProcessingText("C1", function () { PO.updateInventoryReplica(spreadsheet.getActiveSheet()); }, spreadsheet);
            } else if (spreadsheet.getSheetName() == "GCash" && (rg.getA1Notation() == "C1" || rg.getA1Notation() == "I1") && rg.isChecked()) {
                rg.uncheck();
                let sheet = spreadsheet.getActiveSheet();
                //let labelRg = sheet.getRange(rg.getRow(), rg.getColumn()+1)
                //Utils.triggerFuncWithProcessingText(labelRg.getA1Notation(), function() {PO.addGcashToCashReceived(rg, sheet)}, sheet)
                Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { PO.addGcashToCashReceived(rg, sheet); }, sheet);
            } else if (spreadsheet.getSheetName() == "GCash" && rg.getA1Notation() == "L1" && rg.isChecked()) { // Trigger collect gcash on both stores
                rg.uncheck();
                let sheet = spreadsheet.getActiveSheet();
                Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { PO.addGcashToCashReceived(sheet.getRange("C1"), sheet); }, sheet);
                Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { PO.addGcashToCashReceived(sheet.getRange("I1"), sheet); }, sheet);
            } else if (spreadsheet.getSheetName() == "GCash" && (rg.getA1Notation() == "E1" || rg.getA1Notation() == "K1") && rg.isChecked()) {
                rg.uncheck();
                let sheet = spreadsheet.getActiveSheet();
                //let labelRg = sheet.getRange(rg.getRow(), rg.getColumn()+1)
                Utils.triggerFuncWithProcessingText(rg.getA1Notation(), function () { PO.updateInventoryReplica(spreadsheet.getSheetByName("InventoryReplica")); }, sheet);
            } else if (rg.isChecked() && sheetName.startsWith("*")) {
                //Utils.triggerFuncWithProcessingText(a1Not, function() { PO.proxyAddToCashflow(e, a1Not, spreadsheet) }, spreadsheet)
                PO.proxyAddToCashflow(e, a1Not, spreadsheet);
            } else if (spreadsheet.getSheetName().startsWith("Cash flow") && rg.getA1Notation() == "S1" && rg.isChecked()) {
                rg.uncheck();
                Utils.triggerFuncWithProcessingText(a1Not, PO.getUnverifiedSheets, spreadsheet);
            }
        } catch (e) {
            spreadsheet.getRange('U34').setValue(e).setFontColor("red");
            console.log(e.stack);
            Utils.alert(e, "MB PO Err", "", env);
            throw e;
        }
    },



    pullLatestEnding: (storeCode, sheet = SpreadsheetApp.getActive().getActiveSheet(), isGenerateReport = true, env = 'PRD') => {
        console.log("Pulling latest ending for store " + storeCode);
        var inventorySpreadsheet = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode, env));
        var inventorySheets = inventorySpreadsheet.getSheets();
        var latestInvSheet = inventorySheets[inventorySheets.length - 1]
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
            PO.generateReport('', 'M', storeCode, generateReportOffset, env);

        return latestInvSheet;
    },

    nextPO: (storeCode, env = 'PRD') => {
        console.log("Generating next PO for store " + storeCode);
        var spreadsheet = SpreadsheetApp.getActive();

        // Batch read configuration cells with minimal API calls
        var storeCodeFromSheet = spreadsheet.getRange('B6').getValue();
        var oRange = spreadsheet.getRange('O35:P77').getValues();
        var nextPoDateRaw = oRange[0][0]; // O35
        var days2consume = oRange[2][0]; // O37
        var cashFlowAddedText = oRange[42][1]; // P77

        // Check first if PO is registered to cashflow
        if (!cashFlowAddedText || cashFlowAddedText === '') {
            spreadsheet.getRange("U34").setFontColor("red").setFontWeight("bold").setFontStyle("italic").setValue("ERROR: Current PO is not yet added to cash flow");
            SpreadsheetApp.flush();
            Utilities.sleep(5000);
            spreadsheet.getRange("U34").clearContent();
            return;
        }

        var nextPoDate = new Date(nextPoDateRaw);
        nextPoDate.setDate(nextPoDate.getDate() + days2consume);

        const dtFormatted = Utilities.formatDate(nextPoDate, "GMT+8", "MM/dd/yy");
        const prevSheet = spreadsheet.getActiveSheet();
        spreadsheet.duplicateActiveSheet();
        spreadsheet.getActiveSheet().setName('PO D' + dtFormatted + " " + storeCodeFromSheet);

        spreadsheet.getRange('O35').setValue(nextPoDate);

        // Copy to previous order
        spreadsheet.getRange('I31:I').copyTo(spreadsheet.getRange('J31'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

        /* // Cleanup (old)
        spreadsheet.getRange('I31:I').setValue(''); // Prev Orders
        spreadsheet.getRange('O66:O72').setValue(''); // PO/SO Confirmation
        spreadsheet.getRange('P77').setValue(''); // Cash flow added text
        spreadsheet.getRange('P82').setValue(''); // Prev actual PO amount
        spreadsheet.getRange('U63').setValue(''); // Returned SMS
        spreadsheet.getRange('U74').setValue(''); // Returned SMS position
        spreadsheet.getRange('Q65').setValue(''); // Confirmation SMS position
        */

        // Combined Cleanup writes
        spreadsheet.getRangeList(['I31:I', 'O66:O72', 'P77', 'P82', 'U63', 'U74', 'Q65']).clearContent();

        // Hardcode previous projected sales
        const prevSheetProjSalesRg = prevSheet.getRange("G7");
        prevSheetProjSalesRg.setValue(prevSheetProjSalesRg.getValue());

        PO.pullLatestEnding(storeCodeFromSheet, spreadsheet, true, env);
        prevSheet.hideSheet();
    },

    generateReport: (reportName, column, storeCode, rightOffset = 0, env = 'PRD') => {
        console.log("Generating report for store " + storeCode);
        var inventorySpreadsheet = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode, env));
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
        const sheetsToRetain = 30 //inventorySheetsLength-20 //30;
        var archiveSpreadsheet = SpreadsheetApp.openByUrl(Utils.getArchiveInventoryUrl(storeCode, env));
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
            const endRow = PO.getEndRow(sheet);
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
    },

    addPoToCashFlow: (storeCode, env = 'PRD') => {
        var spreadsheet = SpreadsheetApp.getActive();
        var dataRead = spreadsheet.getRange("O35:O66").getValues();
        var confirmationText = dataRead[31][0]; // O66 is 31 rows down from O35
        const dt = Utilities.formatDate(new Date(dataRead[0][0]), "GMT+8", "MM/dd"); // O35 is at index 0

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

        // Append values in batch
        var cashFlowSheet = spreadsheet.getSheetByName(cashFlowSheetName);
        var colValues = cashFlowSheet.getRange("A:A").getValues();
        var count = colValues.filter(String).length;
        cashFlowSheet.getRange(count + 1, 1, 1, 2).setValues([[amt, dt + ' - ' + so]]);

        spreadsheet.getRange('P77').setFontColor("green").setFontWeight("bold").setFontStyle("italic").setValue("Added " + amt + " to Cash flow sheet");
    },

    sendPO: (env = 'PRD') => {
        const smsApiUrl = getSmsApiUrlByConfig(env);
        console.log("Sending PO to " + smsApiUrl);
        var spreadsheet = SpreadsheetApp.getActive();
        var poData = spreadsheet.getRange("O58:U58").getValues()[0];
        var poStr = poData[0]; // O58 is at index 0
        var poNum = poData[6]; // U58 is 6 columns to the right of O58

        var smsApiSheet = SpreadsheetApp.openByUrl(smsApiUrl).getSheetByName("SMS");
        smsApiSheet.appendRow([poNum, poStr, true]);

        var lastRow = smsApiSheet.getLastRow() + 1;
        const importRange = 'IMPORTRANGE("' + smsApiUrl + '", "' + "'SMS'!B\"&U74" + ')';
        console.log(importRange + "\nSMS Row: " + lastRow);

        // Batch write status
        spreadsheet.getRange("U74").setValue(lastRow);
        spreadsheet.getRange("U63").setFormula(importRange);
    },

    confirmPO: (env = 'PRD') => {
        const smsApiUrl = getSmsApiUrlByConfig(env);
        console.log("Confirming PO at " + smsApiUrl);
        var spreadsheet = SpreadsheetApp.getActive();
        var smsRow = spreadsheet.getRange("U74").getValue();
        var smsApiSheet = SpreadsheetApp.openByUrl(smsApiUrl).getSheetByName("SMS");
        smsApiSheet.getRange("'SMS'!C" + smsRow).setFormula("TRUE");

        var lastRow = smsApiSheet.getLastRow() + 1;
        const importRange = 'IMPORTRANGE("' + smsApiUrl + '", "' + "'SMS'!B\"&Q65" + ')';
        console.log(importRange + "\nSMS Row: " + lastRow);

        // Batch write status updates
        spreadsheet.getRangeList(["Q65", "O66"]).getRanges()[0].setValue(lastRow);
        spreadsheet.getRangeList(["Q65", "O66"]).getRanges()[1].setFormula(importRange);
    },

    computeTotalCashCollected: (sheet = SpreadsheetApp.getActiveSheet()) => {
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
    },

    appendToCashReceived: (amt, rg = null, sheet = SpreadsheetApp.getActiveSheet()) => {
        if (!amt) return;
        var colValues = sheet.getRange("F:F").getValues();
        var count = colValues.filter(String).length;

        // Batch write data
        sheet.getRange(count + 2, 5, 1, 2).setValues([[new Date(), amt]]);

        // Print
        if (rg != null) {
            sheet.getRange(rg).setValue("Added " + amt + " to cash received");
            SpreadsheetApp.flush();
            Utilities.sleep(5000);
            sheet.getRange(rg).setValue("");
        }

    },

    appendToExpenses: (startingRow, sheet = SpreadsheetApp.getActiveSheet()) => {
        var rawExpenseSheetName = sheet.getSheetName().replace("Cash flow", "Raw Expenses");
        var rawExpenseSheet = SpreadsheetApp.getActive().getSheetByName(rawExpenseSheetName);

        var expValuesA = sheet.getRange("H:H").getValues();
        var expCount = expValuesA.filter(String).length;
        var receivedValues = sheet.getRange("F:F").getValues();
        var receivedCount = receivedValues.filter(String).length;
        var rawColValues = rawExpenseSheet.getRange("A:A").getValues();
        var rawCount = rawColValues.filter(String).length;

        var expenseInputRange = sheet.getRange("Q" + startingRow + ":R11").getValues();

        let validExpenses = [];
        let cashReceivedRows = [];
        let rawExpenses = [];
        let now = new Date();

        for (let idx = 0; idx < expenseInputRange.length; idx++) {
            var expenseName = expenseInputRange[idx][0];
            var expenseAmount = expenseInputRange[idx][1];
            if (expenseAmount) {
                console.log("Valid expense: " + expenseName);

                // Collect for Expenses (cols G, H, J, N)
                validExpenses.push([now, 0, expenseAmount, expenseName]);

                // Collect for Cash received (cols E, F)
                // Note: using setFormula later or just the value? Original used setFormula("-J" + (expCount + 2))
                // Since we are batching, we should probably record the value or use a relative formula.
                cashReceivedRows.push([now, -expenseAmount]);

                // Collect for Raw Expenses (cols A, B, C, D)
                rawExpenses.push([now, null, expenseName, expenseAmount]);
            }
        }

        if (validExpenses.length > 0) {
            // Write to Expenses (col G is index 7)
            sheet.getRange(expCount + 2, 7, validExpenses.length, 4).setValues(validExpenses.map(v => [v[0], v[1], v[2], v[3]]));

            // Write to Cash received (col E is index 5)
            sheet.getRange(receivedCount + 2, 5, cashReceivedRows.length, 2).setValues(cashReceivedRows);

            // Write to Raw Expenses
            rawExpenseSheet.getRange(rawCount + 1, 1, rawExpenses.length, 4).setValues(rawExpenses);

            // Copy formula for col E in Raw Expenses
            for (let i = 0; i < rawExpenses.length; i++) {
                rawExpenseSheet.getRange("E" + (rawCount + i)).copyTo(rawExpenseSheet.getRange("E" + (rawCount + 1 + i)));
            }
        }
        sheet.getRange("Q5:R11").setValue("");
    },

    appendToCashCollected: (sheet = SpreadsheetApp.getActiveSheet()) => {
        var colValues = sheet.getRange("H:H").getValues();
        var count = colValues.filter(String).length;

        var receivedValues = sheet.getRange("F:F").getValues();
        var receivedCount = receivedValues.filter(String).length;

        // Batch write
        let addedVal = sheet.getRange("O" + (count + 1)).getValue();
        sheet.getRange(receivedCount + 2, 5, 1, 2).setValues([[new Date(), "R" + (count + 1)]]);

        // We need to set the formula for index 6 (col F) specifically if we use setValues with string formula
        sheet.getRange("F" + (receivedCount + 2)).setFormula("R" + (count + 1));

        // Note: R and S columns update
        sheet.getRange(count + 1, 18, 1, 2).setValues([[addedVal, "Added to cash received"]]);
        SpreadsheetApp.flush();
    },

    extractExpensesLoop: (storeCode = "3361", env = 'PRD') => {
        var inventorySpreadSheet = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode, env));
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
            Utils.extractExpenses(dt, employeeName, storeCode, Utils.getEndRow(inventorySheet), inventorySheet, env);
            //console.log("current index: " + i)
            props.setProperty("expenseIndexCounter", i);
            SpreadsheetApp.flush();
        }
    },

    adhocExtractExpensesFromCashFlow: () => {
        var sheet = SpreadsheetApp.getActive().getSheetByName("Cash flow - PCGH");
        var rawExpenseSheetName = sheet.getSheetName().replace("Cash flow", "Raw Expenses");
        var rawExpenseSheet = SpreadsheetApp.getActive().getSheetByName(rawExpenseSheetName);
        var colValuesRaw = rawExpenseSheet.getRange("A:A").getValues();
        var count = colValuesRaw.filter(String).length;

        // Batch read columns G, J, K, N
        // G=7, J=10, K=11, N=14
        var dataRange = sheet.getRange("G2:N393").getValues();
        var batchWrites = [];

        for (var idx = 0; idx < dataRange.length; idx++) {
            var date = dataRange[idx][0];  // Col G
            var expenseName = dataRange[idx][7]; // Col N
            var amount = dataRange[idx][3]; // Col J
            var sales = dataRange[idx][4]; // Col K

            if (!sales) {
                batchWrites.push([date, null, expenseName, amount]);
            }
        }

        if (batchWrites.length > 0) {
            rawExpenseSheet.getRange(count + 1, 1, batchWrites.length, 4).setValues(batchWrites);
        }
    },

    pattyDistribution: (spreadsheet = SpreadsheetApp.getActive(), env = 'PRD') => {
        console.log("Update Patty distribution");
        var lastRow = spreadsheet.getLastRow();
        var allData = spreadsheet.getRange("A1:E" + lastRow).getValues();
        var colBData = allData.map(row => [row[1]]);
        var colCData = allData.map(row => [row[2]]);
        var colEData = allData.map(row => [row[4]]);
        var changed = false;

        Utils.getStoreCodes().forEach((storeCode) => {
            console.log("Current store code: " + storeCode);

            var poSheet = Utils.getLastPoSheet(storeCode, env);
            var poMap = Utils.constructPoMap(poSheet);

            for (var i = 0; i < lastRow; i++) {
                if (allData[i][0] == storeCode) {
                    console.log("[DEBUG] Found store code at row " + (i + 1));

                    for (var j = i + 1; j < lastRow; j++) {
                        var product = allData[j][0];
                        if (product === "") break;

                        if (poMap.has(product)) {
                            colCData[j][0] = poMap.get(product);
                            changed = true;
                        }
                    }

                    // Look for Freezer Top
                    for (var j = i + 1; j < lastRow; j++) {
                        if (allData[j][1] == "Freezer Top") {
                            for (var k = j + 1; k < lastRow; k++) {
                                if (allData[k][1] == "Freezer Bottom") break;
                                if (colBData[k][0] !== "" || colEData[k][0] !== "") {
                                    colBData[k][0] = "";
                                    colEData[k][0] = "";
                                    changed = true;
                                }
                            }
                            break;
                        }
                    }
                }
            }
        });

        if (changed) {
            spreadsheet.getRange("B1:B" + lastRow).setValues(colBData);
            spreadsheet.getRange("C1:C" + lastRow).setValues(colCData);
            spreadsheet.getRange("E1:E" + lastRow).setValues(colEData);
        }

        // Update Inventory Replica
        PO.updateInventoryReplica(spreadsheet.getSheetByName("InventoryReplica"), env);
    },

    updateInventoryReplica: (sheet = SpreadsheetApp.getActive().getActiveSheet(), env = 'PRD') => {
        var sheet = SpreadsheetApp.getActive().getSheetByName("InventoryReplica");
        var lastCol = sheet.getLastColumn();

        // Batch read row 2
        var row2Vals = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
        var updates = [];

        let storeCodes = Utils.getStoreCodes();

        for (var i = 0; i < lastCol; i++) {
            let storeCode = row2Vals[i];
            if (storeCode && storeCode != "" && storeCodes.includes(String(storeCode))) {
                let sheets = SpreadsheetApp.openByUrl(Utils.getInventoryUrl(storeCode, env)).getSheets();
                let lastSheetName = sheets[sheets.length - 1].getSheetName();
                updates.push({ col: i + 2, val: lastSheetName }); // i+2 because we are writing to the column *after* the store code
            }
        }

        // Apply updates. Since columns are not necessarily contiguous, we iterate.
        updates.forEach(u => sheet.getRange(2, u.col).setValue(u.val));
    },

    addGcashToCashReceived: (rg, env = 'PRD') => {
        var gcashSheet = SpreadsheetApp.getActive().getSheetByName("GCash");
        let rgRow = rg.getRow();
        let rgCol = rg.getColumn();

        // Helper functions to get column indices based on rgCol
        const getShiftColIdx = () => rgCol - 2;
        const getGcashColIdx = () => rgCol - 1;
        const getCheckColIdx = () => rgCol + 1;
        const getReplicaColIdx = () => rgCol + 2;
        const getManualColIdx = () => rgCol + 3;

        var replicaCol = getReplicaColIdx();
        var manualCol = getManualColIdx();
        var gcashCol = getGcashColIdx();
        var shiftCol = getShiftColIdx();
        var checkCol = getCheckColIdx();

        // Batch read columns
        var maxRows = 998;
        var startRow = rgRow + 2;
        var replicaVals = gcashSheet.getRange(startRow, replicaCol, maxRows, 1).getValues();
        var manualVals = gcashSheet.getRange(startRow, manualCol, maxRows, 1).getValues();

        var totalGcashRg = gcashSheet.getRange(rgRow, gcashCol);
        var actualGcashRg = gcashSheet.getRange(rgRow, checkCol);

        let totalGcashVal = totalGcashRg.getValue();
        let actualGcashVal = actualGcashRg.getValue();

        if (actualGcashRg.isBlank() || actualGcashVal === "") {
            actualGcashVal = totalGcashVal;
        }

        // Initialize store name
        let storeName = gcashSheet.getRange(rgRow, shiftCol).getValue();
        console.log("Store name: " + storeName);

        // should yield 0 if no over/loss
        let gcashVariance = actualGcashVal - totalGcashVal;
        console.log(`GCash variance: ${gcashVariance}`);
        if (gcashVariance != 0) {
            Utils.cashCollectedAppender(storeName, Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy"), 0, 0, 0, 0, 0, 0, gcashVariance, "GCash", 0);
        }

        // Save actual gcash amount to cash received
        PO.appendToCashReceived(actualGcashVal, null, Utils.getCashFlowSheet(storeName));

        // Processing Replicas and Manuals
        let gcashWrites = new Array(maxRows).fill([""]);
        let checks = new Array(maxRows).fill([false]);
        let replicaManualCounter = 0;

        // Process Replicas
        for (let i = 0; i < maxRows; i++) {
            if (replicaVals[i][0]) {
                gcashWrites[replicaManualCounter] = [replicaVals[i][0]];
                checks[replicaManualCounter] = [true];
                replicaManualCounter++;
            }
        }
        // Process Manuals
        for (let i = 0; i < maxRows; i++) {
            if (manualVals[i][0]) {
                gcashWrites[replicaManualCounter] = [manualVals[i][0]];
                checks[replicaManualCounter] = [true];
                replicaManualCounter++;
            }
        }

        if (replicaManualCounter > 0) {
            // Write collected values back to gcash column starting at startRow
            gcashSheet.getRange(startRow, gcashCol, maxRows, 1).setValues(gcashWrites);
            // Check the boxes for entries
            gcashSheet.getRange(startRow, rgCol, maxRows, 1).setValues(checks);
        }

        // Cleanup original columns
        gcashSheet.getRange(startRow, replicaCol, maxRows, 1).clearContent();
        gcashSheet.getRange(startRow, manualCol, maxRows, 1).clearContent();
        gcashSheet.getRange(startRow, replicaCol, maxRows, 1).uncheck();
        gcashSheet.getRange(startRow, manualCol, maxRows, 1).uncheck();
        actualGcashRg.clearContent();
    },

    getUnverifiedSheets: (sheet = SpreadsheetApp.getActive().getActiveSheet(), env = 'PRD') => {
        let storeCode = sheet.getRange("A1").getValue();
        let inventorySheet = getPoSpreadsheet(getInventoryUrl(storeCode, env), env);
        let unverifiedSheets = showUnverifiedSheets(inventorySheet);
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
    },

    proxyAddToCashflow: (e, a1Not, spreadsheet = SpreadsheetApp.getActive(), env = 'PRD') => {
        let sheet = spreadsheet.getActiveSheet();
        let endRow = Utils.getEndRow(sheet);

        if (a1Not == Utils.getTotalCol() + (endRow + 9)) {
            //sheet.getRange(a1Not).setValue("Processing...")
            sheet.setName(sheet.getSheetName().substring(1));  // Remove the asterisk

            let inventorySpreadsheet = getPoSpreadsheet(getInventoryUrl(getStoreCodeByName(sheet.getRange("A1").getValue()), env), env);
            let inventorySheet = inventorySpreadsheet.getSheetByName(sheet.getSheetName());
            console.log(`Referenced inventory name: ${inventorySheet.getSheetName()}`);

            //let inventoryRg = inventorySheet.getRange(a1Not);
            //inventoryRg.check();   // Check the box for verify in the original inventory sheet
            inventorySheet.getRange(a1Not).check();
            let modifiedEvent = { e, range: inventorySheet.getRange(a1Not), source: inventorySpreadsheet };
            //console.log(e.range.isChecked());

            installedOnEditTriggerInv(modifiedEvent, PropertiesService, env) // Trigger the inventory function;
            spreadsheet.deleteSheet(sheet)  // Delete the replica sheet on PO;
        }
    },

    autoUpdateInventoryReplica: (e, env = 'PRD') => {
        PO.updateInventoryReplica(SpreadsheetApp.getActive().getSheetByName("InventoryReplica"), env);
    }

};