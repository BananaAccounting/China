// Copyright [2015] [Banana.ch SA - Lugano Switzerland]
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
// @id = ch.banana.addon.voucherchinamulticurrency
// @api = 1.0
// @pubdate = 2017-04-24
// @publisher = Banana.ch SA
// @description.en = Voucher multicurrency
// @description.zh = 簿记凭证多货币
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @docproperties = china
// @outputformat = none
// @inputdataform = none
// @timeout = -1








//----------------------------------------
//  Parameters used to create the report
//  Each parameter has:
//      - a unique "id"
//      - an "english" translation
//      - a "chinese" translation
//----------------------------------------
var param = [];
function loadParam() {
    param.push({"id": "showInformation", "english":"Document number not found. Please insert a valid value from the Document column of the Transactions table", "chinese":"未找到文件编号，请您从发生业务表格的文件列中选择一个有效值。"});
    param.push({"id": "printVoucher", "english":"Print Voucher", "chinese":"打印凭证"});
    param.push({"id": "insertDocument", "english":"Insert a document number (or insert * to print all the vouchers)", "chinese":"输入文件编号，或输入'*'号用来打印所有的凭证。"});
    param.push({"id": "voucher", "english":"VOUCHER", "chinese":"记   账   凭   证"});
    param.push({"id": "page", "english":"Page", "chinese":"页面"});
    param.push({"id": "date", "english":"Date", "chinese":"日期"});
    param.push({"id": "year", "english":"Y", "chinese":"年"});
    param.push({"id": "month", "english":"M", "chinese":"月"});
    param.push({"id": "day", "english":"D", "chinese":"日"});
    param.push({"id": "voucherNumber", "english":"Voucher NO.", "chinese":"第   号"});
    param.push({"id": "description", "english":"Description", "chinese":"摘要"});
    param.push({"id": "genLedAc", "english":"GEN. LED. A / C", "chinese":"总账科目"});
    param.push({"id": "subLedAc", "english":"SUB. LED. A / C", "chinese":"明细科目"});
    param.push({"id": "currency", "english":"Currency", "chinese":"币别"});
    param.push({"id": "exchangerate", "english":"Exchange rate", "chinese":"汇率"});
    param.push({"id": "originalCurencyAmount", "english":"Original Amount", "chinese":"原币金额"});
    param.push({"id": "debitAmount", "english":"Debit AMT.", "chinese":"借方金额"});
    param.push({"id": "creditAmount", "english":"Credit AMT.", "chinese":"贷方金额"});
    param.push({"id": "pr", "english":"P.R.", "chinese":"记账"});
    param.push({"id": "attachments", "english":"Attachments", "chinese":"附单据     张"});
    param.push({"id": "total", "english":"Total", "chinese":"合计"});
    param.push({"id": "approved", "english":"Approved", "chinese":"核      准"});
    param.push({"id": "checked", "english":"Checked", "chinese":"复      核"});
    param.push({"id": "entered", "english":"Entered", "chinese":"记      账"});
    param.push({"id": "cashier", "english":"Cashier", "chinese":"出      纳"});
    param.push({"id": "prepared", "english":"Prepared", "chinese":"制      单"});
    param.push({"id": "receiver", "english":"Receiver", "chinese":"签      收"});
}


//-----------------------------------------------------------
//  General parameters that can be changed by the user:
//    - The color of text and table
//    - The number of rows that are printed on each voucher
//-----------------------------------------------------------
var generalParam = {
    "frameColor" : "#32CD32", //chiamarlo framecolor
    "numberOfRows" : 5, // transactions rows
    "PRcolumn" : true,
    "pageSize" : "A4"
    //"pageSize" : "A5 landscape"
}
var currentPage = 1;




/* 
    Main function
*/
function exec(string) {

    var report = Banana.Report.newReport(param.transferVoucher); //Create the report
    var currentTable = Banana.document.cursor.tableName; //Get the name of the current selected table
    var journal = Banana.document.journal(Banana.document.ORIGINTYPE_CURRENT, Banana.document.ACCOUNTTYPE_NORMAL); //Create the Journal
    var docList = getDocList(); //Create a list with all the doc numbers    
    
    loadParam(); //Load all the parameters
    
    if (currentTable === "Transactions") { //If the current table is the Transactions table, we take the value of the column "Doc" of the selected row
        var transactions = Banana.document.table('Transactions');
        
        try {
            var tDoc = transactions.row(Banana.document.cursor.rowNr).value('Doc');
        } catch(error) {}

        /* Open dialog window asking to insert a voucher number (or * for all vouchers):
        If transactions table is selected, we use the value of the selected row as default
        If transactions table is not selected, we don't use any default value */
        var docNumber = Banana.Ui.getText(getValue(param, "printVoucher", "chinese"), getValue(param, "insertDocument", "chinese"), tDoc);
    } else {
        var docNumber = Banana.Ui.getText(getValue(param, "printVoucher", "chinese"), getValue(param, "insertDocument", "chinese"), '');
    }

    if (docNumber && docNumber !== "*") { //The user insert a correct voucher value => prints single voucher
        if (docList.indexOf(docNumber) > -1) { //Check if the docNumber exists in the docList array
            docList = [];
            docList.push(docNumber);
            printVoucher(report, docList, journal);
        }
        else { //docNumber doesn't exists
            Banana.Ui.showInformation("", getValue(param, "showInformation", "chinese"));
            return;
        }
    }
    else if (docNumber === "*") { //The user insert the "*" => prints all the vouchers
        printVoucher(report, docList, journal);
    }
    else { //User doesn't insert anything then clic "ok", or clic on "cancel" button
        return;
    }

    //Add the styles
    var stylesheet = createStyleSheet();
    Banana.Report.preview(report, stylesheet);
}


/*
    Function that creates an array of rows and use it to print the voucher report
*/
function printVoucher(report, listDoc, journal) {

    var rowsToProcess = []; //Array to store all the rows that will be processed

    for (var j = 0; j < listDoc.length; j++) {
        for (var i = 0; i < journal.rowCount; i++) {
            var tRow = journal.row(i);
            if (listDoc[j] && tRow.value('Doc') === listDoc[j]) { //Check if the actual Doc number of the list is equal to the Doc value of the journal's row
                rowsToProcess.push(i); //Add the row to the array
                if (rowsToProcess.length == generalParam.numberOfRows) { //Check if there are N elements into the array. We want a maximum of N rows per voucher
                    createVoucherReport(journal, report, listDoc[j], rowsToProcess); //Function call to create the voucher for the given rows
                    rowsToProcess = []; //Array reset
                }
            }
        }
        if (rowsToProcess.length > 0) { //Check if the array contains something
            createVoucherReport(journal, report, listDoc[j], rowsToProcess); //Function call to create the voucher for the given rows
            rowsToProcess = []; //Array reset
        }
        currentPage = 1;
    }
}


/*
    Function that creates the voucher for the given Doc number
*/
function createVoucherReport(journal, report, docNumber, rowsToProcess) {

    //Each voucher on a new page
    report.addPageBreak();

    var date = "";
    var doc = ""; 
    var desc = "";
    var totDebit = "";
    var totCredit = "";
    var line = 0;
    var totTransaction = calculateTotalTransaction(docNumber, journal, report); //Calculate the total value of the transaction

    /* Calculate the total number of pages used to print each voucher */
    var totRows = calculateTotalRowsOfTransaction(docNumber, journal, report);
    if (totRows > generalParam.numberOfRows) {
        var totalPages = Banana.SDecimal.divide(totRows, generalParam.numberOfRows);
        totalPages = parseFloat(totalPages);
        totalPages += 0.499999999999999;
        var context = {'decimals' : 0, 'mode' : Banana.SDecimal.HALF_UP}
        totalPages = Banana.SDecimal.round(totalPages, context);
    } else {
        var totalPages = 1;
    }

    //Define table info: name and columns
    var tableInfo = report.addTable("table_info");
    var col1 = tableInfo.addColumn("col1");
    var col2 = tableInfo.addColumn("col2");
    var col3 = tableInfo.addColumn("col3");
    var col4 = tableInfo.addColumn("col4");
    var col5 = tableInfo.addColumn("col5");
    var col6 = tableInfo.addColumn("col6");
    var col7 = tableInfo.addColumn("col7");
    var col8 = tableInfo.addColumn("col8");
    var col9 = tableInfo.addColumn("col9");
    var col10 = tableInfo.addColumn("col10");
    var col11 = tableInfo.addColumn("col11");

    //Define table transactions: name and columns

    if (Banana.document.table('ExchangeRates')) {
        var table = report.addTable("table");
        var ct1 = table.addColumn("ct1");
        var ct2 = table.addColumn("ct2");
        var ct3 = table.addColumn("ct3");
        var ct4 = table.addColumn("ct4");
        var ct5 = table.addColumn("ct5");
        var ct5 = table.addColumn("ct6");
        var cts1 = table.addColumn("cts1");
        var ct4 = table.addColumn("ct7");
        var cts2 = table.addColumn("cts2");
        var ct5 = table.addColumn("ct8");
        var cts3 = table.addColumn("cts3");
        var ct6 = table.addColumn("ct9");
    } else {
        var table = report.addTable("table");
        var c1 = table.addColumn("c1");
        var c2 = table.addColumn("c2");
        var c3 = table.addColumn("c3");
        var cs1 = table.addColumn("cs1");
        var c4 = table.addColumn("c4");
        var cs2 = table.addColumn("cs2");
        var c5 = table.addColumn("c5");
        var cs3 = table.addColumn("cs3");
        var c6 = table.addColumn("c6");
    }

    //Define table signatures: name and columns
    var tableSignatures = report.addTable("table_signatures");
    var colSig1 = tableSignatures.addColumn("colSig1");
    var colSig2 = tableSignatures.addColumn("colSig2");
    var colSig3 = tableSignatures.addColumn("colSig3");
    var colSig4 = tableSignatures.addColumn("colSig4");
    var colSig5 = tableSignatures.addColumn("colSig5");
    var colSig6 = tableSignatures.addColumn("colSig6");

    //Select from the journal all the given rows, take all the needed data, calculate totals and create the vouchers print
    for (j = 0; j < rowsToProcess.length; j++) {
        for (i = 0; i < journal.rowCount; i++) {
            if (i == rowsToProcess[j]) {
                var tRow = journal.row(i);
                date = tRow.value("Date");
            }
        }
    }

    /* 1. Print the date and the voucher number */
    printInfoVoucher(tableInfo, report, docNumber, date, totalPages);

    /* 2. Print the all the transactions data */
    printTransactions(table, journal, line, rowsToProcess);
    
    /* 3. Print the total line */
    printTotal(table, totDebit, totCredit, report, totalPages, totTransaction);
    
    /* 4. Print the signatures line */
    printSignatures(tableSignatures, report);

    currentPage++;
}


/* Function that prints the date and the voucher number */
function printInfoVoucher(tableInfo, report, docNumber, date, totalPages) {

    //Title
    tableRow = tableInfo.addRow();
    tableRow.addCell(getValue(param, "voucher", "chinese"), "heading1 alignCenter", 11);

    tableRow = tableInfo.addRow();
    tableRow.addCell(getValue(param, "voucher", "english"), "heading2 alignCenter", 11);


    //Date and Voucher number
    var d = Banana.Converter.toDate(date);
    var year = d.getFullYear();
    var month = d.getMonth()+1;
    var day = d.getDate();

    tableRow = tableInfo.addRow();
    
    var cell1 = tableRow.addCell(getValue(param, "page", "chinese") + " " +  currentPage + " / " + totalPages, "", 1);
    cell1.addParagraph(getValue(param, "page", "english"), "");

    var cell2 = tableRow.addCell("", "alignRight border-top-double", 1);
    cell2.addParagraph(getValue(param, "date", "chinese") + ": ");

    var cell3 = tableRow.addCell("", "alignCenter border-top-double", 1);
    cell3.addParagraph(year, "text-black");

    var cell4 = tableRow.addCell("", "alignCenter border-top-double", 1);
    cell4.addParagraph(getValue(param, "year", "chinese"), "");
    cell4.addParagraph(getValue(param, "year", "english"), "");
    
    var cell5 = tableRow.addCell("", "alignCenter border-top-double", 1);
    cell5.addParagraph(month, "text-black");

    var cell6 = tableRow.addCell("", "alignCenter border-top-double", 1);
    cell6.addParagraph(getValue(param, "month", "chinese"), "");
    cell6.addParagraph(getValue(param, "month", "english"), "");

    var cell7 = tableRow.addCell("", "alignCenter border-top-double", 1);
    cell7.addParagraph(day, "text-black");

    var cell8 = tableRow.addCell("", "alignCenter border-top-double", 1);
    cell8.addParagraph(getValue(param, "day", "chinese"), "");
    cell8.addParagraph(getValue(param, "day", "english"), "");

    var cell9 = tableRow.addCell("", "", 1);    
    cell9.addParagraph(" ", "");

    var cell10 = tableRow.addCell("", "padding-left alignCenter", 1);
    cell10.addParagraph(getValue(param, "voucherNumber", "chinese"), "");
    cell10.addParagraph(getValue(param, "voucherNumber", "english"), "");
    
    var cell11 = tableRow.addCell("", "padding-left", 1);
    cell11.addParagraph(" ", "");
    cell11.addParagraph(docNumber, "text-black");

    report.addParagraph(" ", "");
}


/* Function that prints all the transactions data */
function printTransactions(table, journal, line, rowsToProcess) {

    /* Table header */
    var strPr = getValue(param, "pr", "chinese");
    var res = strPr.split("");

    var tableHeader = table.getHeader();
    tableRow = tableHeader.addRow();

    if (Banana.document.table('ExchangeRates')) {
        tableRow.addCell(getValue(param, "description", "chinese"), "alignCenter border-left-black border-top-black border-right", 1);
        tableRow.addCell(getValue(param, "genLedAc", "chinese"), "alignCenter border-left border-top-black border-right", 1);
        tableRow.addCell(getValue(param, "subLedAc", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell(getValue(param, "currency", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell(getValue(param, "exchangerate", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell(getValue(param, "originalCurencyAmount", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
        tableRow.addCell(getValue(param, "debitAmount", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
        tableRow.addCell(getValue(param, "creditAmount", "chinese"), "alignCenter border-left border-top-black border-right-black",1);
        tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
        if (generalParam.PRcolumn) {
            tableRow.addCell(getValue(param, "pr", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        }

        tableRow = tableHeader.addRow();
        tableRow.addCell(getValue(param, "description", "english"), "alignCenter border-left-black border-right border-bottom-black", 1);
        tableRow.addCell(getValue(param, "genLedAc", "english"), "alignCenter border-left border-right border-bottom-black", 1);
        tableRow.addCell(getValue(param, "subLedAc", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "currency", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "exchangerate", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "originalCurencyAmount", "english"), "alignCenter border-left border-right-black", 1);
        tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "debitAmount", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "creditAmount", "english"), "alignCenter border-left border-right-black border-bottom-black",1);
        tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);
        if (generalParam.PRcolumn) {
            tableRow.addCell(getValue(param, "pr", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        }


        /* Print transactions rows */
        var tmpDescription;
        for (var j = 0; j < rowsToProcess.length; j++) {
            
            for (var i = 0; i < journal.rowCount; i++) {
                
                if (i == rowsToProcess[j]) {
                    var tRow = journal.row(i);
                    var amount = tRow.value('JAmount');
                    
                    //var amountAccountCurrency = Banana.SDecimal.abs(tRow.value("JAmountAccountCurrency"));
                    var amountAccountCurrency = Banana.SDecimal.abs(tRow.value("JAmountAccountCurrency"));

                    //var accountCurrency = tRow.value('JAccountCurrency'); //CHF, EUR...
                    var exchangeCurrency = tRow.value('ExchangeCurrency'); //CHF, EUR
                    

                    var exchangeRate = Banana.SDecimal.round(tRow.value('ExchangeRate'), {'decimals':4});

                    // Banana.console.log(amount + ", " + amountAccountCurrency);
                    // var exchangeRate = Banana.SDecimal.divide(amount, amountAccountCurrency, {'decimals':4});
                    // if (exchangeRate === "1.0000") {
                    //     exchangeRate = "1";
                    // }


                    //Banana.console.log("=>" + tRow.value("JAmountTransactionCurrency"));
                    //JAccountCurrency
                    //JAmountAccountCurrency
                    // JTransactionCurrency
                    // JAmountTransactionCurrency
                    // JTransactionCurrencyConversionRate
                    // JDebitAmountAccountCurrency
                    // JCreditAmountAccountCurrency
                    // JBalanceAccountCurrency

                    
                    //We check if the current description is equal to the previous (same transaction)
                    //In this case we print the description only in the first row
                    tableRow = table.addRow();
                    if (tmpDescription !== tRow.value('Description')) {
                        tmpDescription = tRow.value('Description');
                        tableRow.addCell(tRow.value('Description'), "text-black padding-left border-left-black border-top border-right border-bottom", 1);
                    } else {
                        tableRow.addCell("", "text-black padding-left border-left-black border-top border-right border-bottom", 1);
                    }
                    tableRow.addCell(tRow.value('JAccountDescription'), "text-black padding-left border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-left border-left border-top border-right-black border-bottom", 1);
                    tableRow.addCell(exchangeCurrency, "text-black padding-left border-left border-top border-right-black border-bottom", 1);
                    tableRow.addCell(exchangeRate, "text-black padding-left border-left border-top border-right-black border-bottom", 1);
                    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amountAccountCurrency), "text-black padding-left border-left border-top border-right-black border-bottom", 1);
                    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    
                    if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ) { // Debit
                        tableRow.addCell("¥ " + Banana.Converter.toLocaleNumberFormat(amount), "text-black padding-right border-left-black border-top border-right-black border-bottom alignRight", 1);
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    } else { // Credit
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                        tableRow.addCell("¥ " + Banana.Converter.toLocaleNumberFormat(amount), "text-black padding-right border-left-black border-top border-right-black border-bottom alignRight", 1);
                    }

                    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    if (generalParam.PRcolumn) {
                        tableRow.addCell(" ", "text-black alignCenter border-left border-top border-right-black border-bottom", 1);
                    }
                    line++;
                }
            }
        }

        //Add empty lines if current lines are not enough
        while (line < generalParam.numberOfRows) {
            tableRow = table.addRow();
            tableRow.addCell("", "border-left-black border-top border-right border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right-black border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right-black border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right-black border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right-black border-bottom", 1);
            tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
            tableRow.addCell("**********", "text-black padding-right border-left border-top border-right-black border-bottom alignCenter", 1);
            tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
            tableRow.addCell("**********", "text-black padding-right border-left border-top border-right-black border-bottom alignCenter", 1);
            tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
            if (generalParam.PRcolumn) {
                tableRow.addCell(" ", "text-black alignCenter border-left border-top border-right-black border-bottom", 1);
            }
            line++;
        }
    } else {

        tableRow.addCell(getValue(param, "description", "chinese"), "alignCenter border-left-black border-top-black border-right", 1);
        tableRow.addCell(getValue(param, "genLedAc", "chinese"), "alignCenter border-left border-top-black border-right", 1);
        tableRow.addCell(getValue(param, "subLedAc", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
        tableRow.addCell(getValue(param, "debitAmount", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
        tableRow.addCell(getValue(param, "creditAmount", "chinese"), "alignCenter border-left border-top-black border-right-black",1);
        tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
        if (generalParam.PRcolumn) {
            tableRow.addCell(getValue(param, "pr", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
        }

        tableRow = tableHeader.addRow();
        tableRow.addCell(getValue(param, "description", "english"), "alignCenter border-left-black border-right border-bottom-black", 1);
        tableRow.addCell(getValue(param, "genLedAc", "english"), "alignCenter border-left border-right border-bottom-black", 1);
        tableRow.addCell(getValue(param, "subLedAc", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "debitAmount", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);
        tableRow.addCell(getValue(param, "creditAmount", "english"), "alignCenter border-left border-right-black  border-bottom-black",1);
        tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);
        if (generalParam.PRcolumn) {
            tableRow.addCell(getValue(param, "pr", "english"), "alignCenter border-left border-right-black border-bottom-black", 1);
        }

        /* Print transactions rows */
        var tmpDescription;
        for (var j = 0; j < rowsToProcess.length; j++) {
            
            for (var i = 0; i < journal.rowCount; i++) {
                
                if (i == rowsToProcess[j]) {
                    var tRow = journal.row(i);
                    var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
                    
                    //We check if the current description is equal to the previous (same transaction)
                    //In this case we print the description only in the first row
                    tableRow = table.addRow();
                    if (tmpDescription !== tRow.value('Description')) {
                        tmpDescription = tRow.value('Description');
                        tableRow.addCell(tRow.value('Description'), "text-black padding-left border-left-black border-top border-right border-bottom", 1);
                    } else {
                        tableRow.addCell("", "text-black padding-left border-left-black border-top border-right border-bottom", 1);
                    }
                    tableRow.addCell(tRow.value('JAccountDescription'), "text-black padding-left border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-left border-left border-top border-right-black border-bottom", 1);
                    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    
                    if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ) { // Debit
                        tableRow.addCell("¥ " + Banana.Converter.toLocaleNumberFormat(amount), "text-black padding-right border-left-black border-top border-right-black border-bottom alignRight", 1);
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    } else { // Credit
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                        tableRow.addCell("¥ " + Banana.Converter.toLocaleNumberFormat(amount), "text-black padding-right border-left-black border-top border-right-black border-bottom alignRight", 1);
                    }

                    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    if (generalParam.PRcolumn) {
                        tableRow.addCell(" ", "text-black alignCenter border-left border-top border-right-black border-bottom", 1);
                    }
                    line++;
                }
            }
        }

        //Add empty lines if current lines are not enough
        while (line < generalParam.numberOfRows) {
            tableRow = table.addRow();
            tableRow.addCell("", "border-left-black border-top border-right border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right border-bottom", 1);
            tableRow.addCell("", "border-left border-top border-right-black border-bottom", 1);
            tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
            tableRow.addCell("**********", "text-black padding-right border-left border-top border-right-black border-bottom alignCenter", 1);
            tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
            tableRow.addCell("**********", "text-black padding-right border-left border-top border-right-black border-bottom alignCenter", 1);
            tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
            if (generalParam.PRcolumn) {
                tableRow.addCell(" ", "text-black alignCenter border-left border-top border-right-black border-bottom", 1);
            }
            line++;
        }
    }
}








/* Function that prints the total line of the voucher */
function printTotal(table, totDebit, totCredit, report, totalPages, totTransaction) {

    //Attachments and Total texts
    tableRow = table.addRow();

    if (Banana.document.table('ExchangeRates')) {
        tableRow.addCell(getValue(param, "attachments", "chinese") + "   " + getValue(param, "attachments", "english"), "alignCenter padding-left border-left-black border-top-black border-right border-bottom-black", 3);
        tableRow.addCell(getValue(param, "total", "chinese") + "   " + getValue(param, "total", "english"), "alignCenter padding-left border-left border-top-black border-right-black border-bottom-black", 3);
    } else {
        tableRow.addCell(getValue(param, "attachments", "chinese") + "   " + getValue(param, "attachments", "english"), "alignCenter padding-left border-left-black border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(getValue(param, "total", "chinese") + "   " + getValue(param, "total", "english"), "alignCenter padding-left border-left border-top-black border-right-black border-bottom-black", 2);
    }

    //Total value
    if (currentPage != totalPages) {
        var totDebit = "";
        var totCredit = "";
    } else {
        var totDebit = "¥ " + Banana.Converter.toLocaleNumberFormat(totTransaction[0]);
        var totCredit = "¥ " + Banana.Converter.toLocaleNumberFormat(totTransaction[1]);
    }

    tableRow.addCell("", "border-top-black border-left-black border-right-black border-bottom-black", 1);
    tableRow.addCell(totDebit, "text-black padding-right border-left-black border-top-black border-right-black border-bottom-black alignRight", 1);
    tableRow.addCell("", "border-top-black border-left-black border-right-black border-bottom-black", 1);
    tableRow.addCell(totCredit, "text-black padding-right border-left-black border-top-black border-right-black border-bottom-black alignRight", 1);
    tableRow.addCell("", "border-left-black border-top-black border-right-black border-bottom-black", 1);
    if (generalParam.PRcolumn) {
        tableRow.addCell(" ", "text-black alignCenter border-left border-top-black border-right-black border-bottom-black", 1);
    }
}





/* Function that prints the signature part of the voucher */
function printSignatures(table, report) {

    tableRow = table.addRow();
    var cellApproved = tableRow.addCell("", "", 1);
    cellApproved.addParagraph(getValue(param, "approved", "chinese") + ": ", "");
    cellApproved.addParagraph(getValue(param, "approved", "english"), "");
    
    var cellChecked = tableRow.addCell("", "", 1);
    cellChecked.addParagraph(getValue(param, "checked", "chinese") + ": ", "");
    cellChecked.addParagraph(getValue(param, "checked", "english"), "");

    var cellEntered = tableRow.addCell("", "", 1);
    cellEntered.addParagraph(getValue(param, "entered", "chinese") + ": ", "");
    cellEntered.addParagraph(getValue(param, "entered", "english"), "");

    var cellCashier = tableRow.addCell("", "", 1);
    cellCashier.addParagraph(getValue(param, "cashier", "chinese") + ": ", "");
    cellCashier.addParagraph(getValue(param, "cashier", "english"), "");

    var cellPrepared = tableRow.addCell("", "", 1);
    cellPrepared.addParagraph(getValue(param, "prepared", "chinese") + ": ", "");
    cellPrepared.addParagraph(getValue(param, "prepared", "english"), "");

    var cellReceiver = tableRow.addCell("", "", 1);
    cellReceiver.addParagraph(getValue(param, "receiver", "chinese") + ": ", "");
    cellReceiver.addParagraph(getValue(param, "receiver", "english"), "");
}




//Function that calculates the total rows of the transaction
function calculateTotalRowsOfTransaction(docNumber, journal, report) {
    var totRows = 0;
    for (var i = 0; i < journal.rowCount; i++) {
        var tRow = journal.row(i);
        if (tRow.value('Doc') === docNumber) {
            totRows++;
        }
    }
    return totRows;
}


//Function that calculates the total of the debit/credit of the transaction for a given docNumber
function calculateTotalTransaction(docNumber, journal, report) {
   
    var tDebit = "";
    var tCredit = "";
    var arr = [];

    for (var i = 0; i < journal.rowCount; i++) {
        var tRow = journal.row(i);
        if (tRow.value('Doc') === docNumber) {
            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ) {
                tDebit = Banana.SDecimal.add(tDebit, amount, {'decimals':2});
            } else {
                tCredit = Banana.SDecimal.add(tCredit, amount, {'decimals':2});
            }
        }
    }
    arr.push(tDebit);
    arr.push(tCredit);
    return arr;
}





/* Function that creates a list with all the values taken from the columna "Doc" of the table "Transactions" */
function getDocList() {
    var str = [];
    for (var i = 0; i < Banana.document.table('Transactions').rowCount; i++) {
        var tRow = Banana.document.table('Transactions').row(i);
        if (tRow.value("Doc")) {
            str.push(tRow.value("Doc"));
        }
    }

    //Removing duplicates
    for (var i = 0; i < str.length; i++) {
        for (var x = i+1; x < str.length; x++) {
            if (str[x] === str[i]) {
                str.splice(x,1);
                --x;
            }
        }
    }

    return str;
}


/* Function that returns the whole object */
function getObject(form, id) {
    for (var i = 0; i < form.length; i++) {
        if (form[i]["id"] === id) {
            return form[i];
        }
    }
    Banana.document.addMessage("Couldn't find object with id: " + id);
}


/* Function that returns a specific field value from the object */
function getValue(form, id, field) {
    var searchId = id.trim();
    for (var i = 0; i < form.length; i++) {
        if (form[i].id === searchId) {
            return form[i][field];
        }
    }
    Banana.document.addMessage("Couldn't find object with id:" + id);
}


/* Function that creates all the styles used to print the report */
function createStyleSheet() {
    
    //Create stylesheet
    var stylesheet = Banana.Report.newStyleSheet();
    
    //Set page layout
    var pageStyle = stylesheet.addStyle("@page");

    //Set the margins
    pageStyle.setAttribute("margin", "10mm 5mm 10mm 5mm");
    pageStyle.setAttribute("size", generalParam.pageSize);


    /*
        General styles
    */
    stylesheet.addStyle("body", "font-family : Times New Roman; font-size:10pt; color:" + generalParam.frameColor);
    stylesheet.addStyle(".font-size-digits", "font-size:6pt");
    stylesheet.addStyle(".text-green", "color:" + generalParam.frameColor);
    stylesheet.addStyle(".text-black", "color:black");

    stylesheet.addStyle(".border-top", "border-top: thin solid " +  generalParam.frameColor);
    stylesheet.addStyle(".border-right", "border-right: thin solid " +  generalParam.frameColor);
    stylesheet.addStyle(".border-bottom", "border-bottom: thin solid " +  generalParam.frameColor);
    stylesheet.addStyle(".border-left", "border-left: thin solid " +  generalParam.frameColor);
    stylesheet.addStyle(".border-left-1px", "border-left: 1px solid " +  generalParam.frameColor);
    stylesheet.addStyle(".border-top-double", "border-top: 0.8px double black");

    stylesheet.addStyle(".border-top-black", "border-top: 0.5px solid black");
    stylesheet.addStyle(".border-right-black", "border-right: thin solid black");
    stylesheet.addStyle(".border-bottom-black", "border-bottom: thin solid black");
    stylesheet.addStyle(".border-left-black", "border-left: thin solid black");

    stylesheet.addStyle(".padding-left", "padding-left:5px");
    stylesheet.addStyle(".padding-right", "padding-right:5px");
    stylesheet.addStyle(".underLine", "border-top:thin double black");

    stylesheet.addStyle(".heading1", "font-size:16px;font-weight:bold");
    stylesheet.addStyle(".heading2", "font-size:12px");
    stylesheet.addStyle(".bold", "font-weight:bold");
    stylesheet.addStyle(".alignRight", "text-align:right");
    stylesheet.addStyle(".alignCenter", "text-align:center");

 
    /* 
        Info table style
    */
    style = stylesheet.addStyle("table_info");
    style.setAttribute("width", "100%");
    style.setAttribute("font-size", "8px");
    stylesheet.addStyle("table.table_info td", "padding-bottom: 2px; padding-top: 5px;");

    //Columns for the info table
    stylesheet.addStyle(".col1", "width:37%");
    stylesheet.addStyle(".col2", "width:5%");
    stylesheet.addStyle(".col3", "width:5%");
    stylesheet.addStyle(".col4", "width:2%");
    stylesheet.addStyle(".col5", "width:5%");
    stylesheet.addStyle(".col6", "width:2%");
    stylesheet.addStyle(".col7", "width:5%");
    stylesheet.addStyle(".col8", "width:2%");
    stylesheet.addStyle(".col9", "width:15%");
    stylesheet.addStyle(".col10", "width:15%");
    stylesheet.addStyle(".col11", "width:%");

    /*
        Transactions table style
    */
    style = stylesheet.addStyle("table");
    style.setAttribute("width", "100%");
    style.setAttribute("font-size", "8px");
    stylesheet.addStyle("table.table td", "padding-bottom: 4px; padding-top: 6px");

    //Columns for the transactions table
    stylesheet.addStyle(".c1", "width:25%");
    stylesheet.addStyle(".c2", "width:25%");
    stylesheet.addStyle(".c3", "width:25%");
    stylesheet.addStyle(".cs1", "width:0.4%");
    stylesheet.addStyle(".c4", "width:10%");
    stylesheet.addStyle(".cs2", "width:0.4%");
    stylesheet.addStyle(".c5", "width:10%");
    stylesheet.addStyle(".cs3", "width:0.4%");
    stylesheet.addStyle(".c6", "width:3.2%");

    stylesheet.addStyle(".ct1", "width:20%");
    stylesheet.addStyle(".ct2", "width:20%");
    stylesheet.addStyle(".ct3", "width:10%");
    stylesheet.addStyle(".ct4", "width:10%");
    stylesheet.addStyle(".ct5", "width:10%");
    stylesheet.addStyle(".ct6", "width:10%");
    stylesheet.addStyle(".cts1", "width:0.4%");
    stylesheet.addStyle(".ct7", "width:10%");
    stylesheet.addStyle(".cts2", "width:0.4%");
    stylesheet.addStyle(".ct8", "width:10%");
    stylesheet.addStyle(".cts3", "width:0.4%");
    stylesheet.addStyle(".ct9", "width:3.2%");

    /*
        Signatures table style
    */
    style = stylesheet.addStyle("table_signatures");
    style.setAttribute("width", "100%");
    style.setAttribute("font-size", "8px");
    stylesheet.addStyle("table.table_signatures td", "padding-bottom: 5px; padding-top: 5px;");

    //Column for the signatures table
    stylesheet.addStyle(".colSig1", "width:16.6%");
    stylesheet.addStyle(".colSig2", "width:16.6%");
    stylesheet.addStyle(".colSig3", "width:16.6%");
    stylesheet.addStyle(".colSig4", "width:16.6%");
    stylesheet.addStyle(".colSig5", "width:16.6%");
    stylesheet.addStyle(".colSig6", "width:16.6%");

    return stylesheet;
}


