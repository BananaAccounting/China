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
// @id = ch.banana.addon.voucherchina
// @api = 1.0
// @pubdate = 2017-04-24
// @publisher = Banana.ch SA
// @description.en = Voucher
// @description.cn = 记账凭证
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
    param.push({"id": "date", "english":"Date", "chinese":"日期"});
    param.push({"id": "year", "english":"Y", "chinese":"年"});
    param.push({"id": "month", "english":"M", "chinese":"月"});
    param.push({"id": "day", "english":"D", "chinese":"日"});
    param.push({"id": "voucherNumber", "english":"Voucher NO.", "chinese":"第   号"});
    param.push({"id": "description", "english":"Description", "chinese":"摘要"});
    param.push({"id": "genLedAc", "english":"GEN. LED. A / C", "chinese":"总账科目"});
    param.push({"id": "subLedAc", "english":"SUB. LED. A / C", "chinese":"明细科目"});
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
    "textColor" : "#32CD32",
    // transactions rows
    "numberOfRows" : 6,
    "pageSize" : "A4"
    //"pageSize" : "A5 landscape"
}





/* 
    Main function
*/
function exec(string) {

    //Create the report
    var report = Banana.Report.newReport(param.transferVoucher);
    
    //Load all the parameters
    loadParam();
    
    //Get the name of the current selected table
    var currentTable = Banana.document.cursor.tableName;

    //If the current table is the Transactions table, we take the value of the column "Doc" of the selected row
    if (currentTable === "Transactions") {
        var transactions = Banana.document.table('Transactions');
        var tDoc = transactions.row(Banana.document.cursor.rowNr).value('Doc');

        //Open dialog window asking to insert a voucher number (or * for all vouchers):
        //If transactions table is selected, we use the value of the selected row as default
        var docNumber = Banana.Ui.getText(getValue(param, "printVoucher", "chinese"), getValue(param, "insertDocument", "chinese"), tDoc);
    } else {
        //Open dialog window asking to insert a voucher number (or * for all vouchers):
        //If transactions table is not selected, we don't use any default value
        var docNumber = Banana.Ui.getText(getValue(param, "printVoucher", "chinese"), getValue(param, "insertDocument", "chinese"), '');
    }

    //Create the Journal table which contains all the data of the accounting
    var journal = Banana.document.journal(Banana.document.ORIGINTYPE_CURRENT, Banana.document.ACCOUNTTYPE_NORMAL);
    
    //Create a list with all the doc numbers
    var docList = getDocList();

    //The user insert a correct voucher value => prints single voucher
    if (docNumber && docNumber !== "*") {

        //Check if the docNumber exists in the docList array
        if (docList.indexOf(docNumber) > -1) {
            docList = [];
            docList.push(docNumber);
            printVoucher(report, docList, journal);
        }
        else { //docNumber doesn't exists
            Banana.Ui.showInformation("", getValue(param, "showInformation", "chinese"));
            return;
        }
    }
    //The user insert the "*" => prints all the vouchers
    else if (docNumber === "*") {
        printVoucher(report, docList, journal);

    }
    //User doesn't insert anything then clic "ok", or clic on "cancel" button
    else {
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

    //Array to store all the rows that will be processed
    var rowsToProcess = []; 

    for (var j = 0; j < listDoc.length; j++)
    {
        var docNumber = listDoc[j];

        for (var i = 0; i < journal.rowCount; i++) 
        {
            var tRow = journal.row(i);

            //Check if the actual Doc number of the list is equal to the Doc value of the journal's row
            if (docNumber && tRow.value('Doc') === docNumber) 
            {
                //Add the row to the array
                rowsToProcess.push(i);

                //Check if there are six elements (six rows) into the array.
                //We want a maximum of six rows per voucher
                if (rowsToProcess.length == generalParam.numberOfRows)
                {
                    //Function call to create the voucher for the given rows
                    createVoucherReport(journal, report, docNumber, rowsToProcess);

                    //Array reset
                    rowsToProcess = [];
                }
            }
        }
    
        //Check if the array contains something
        if (rowsToProcess.length > 0) 
        {
            //Function call to create the voucher for the given rows
            createVoucherReport(journal, report, docNumber, rowsToProcess);

            //Array reset
            rowsToProcess = [];
        }
    }
}


/*
    Function that creates the voucher for the given Doc number
*/
function createVoucherReport(journal, report, docNumber, rowsToProcess) {

    //Each voucher on a new page
    report.addPageBreak();

    //Define some variables
    var date = "";
    var doc = ""; 
    var desc = "";
    var totDebit = "";
    var totCredit = "";
    var line = 0;

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
    var table = report.addTable("table");
    var c1 = table.addColumn("c1");
    var c2 = table.addColumn("c2");
    var c3 = table.addColumn("c3");
    var cs1 = table.addColumn("cs1");
    var c4 = table.addColumn("c4");
    var c5 = table.addColumn("c5");
    var c6 = table.addColumn("c6");
    var c7 = table.addColumn("c7");
    var c8 = table.addColumn("c8");
    var c9 = table.addColumn("c9");
    var c10 = table.addColumn("c10");
    var c11 = table.addColumn("c11");
    var c12 = table.addColumn("c12");
    var c13 = table.addColumn("c13");
    var c14 = table.addColumn("c14");
    var cs2 = table.addColumn("cs2");
    var c15 = table.addColumn("c15");
    var c16 = table.addColumn("c16");
    var c17 = table.addColumn("c17");
    var c18 = table.addColumn("c18");
    var c19 = table.addColumn("c19");
    var c20 = table.addColumn("c20");
    var c21 = table.addColumn("c21");
    var c22 = table.addColumn("c22");
    var c23 = table.addColumn("c23");
    var c24 = table.addColumn("c24");
    var c25 = table.addColumn("c25");
    var cs3 = table.addColumn("cs3");
    var c26 = table.addColumn("c26");

    //Define table signatures: name and columns
    var tableSignatures = report.addTable("table_signatures");
    var colSig1 = tableSignatures.addColumn("colSig1");
    var colSig2 = tableSignatures.addColumn("colSig2");
    var colSig3 = tableSignatures.addColumn("colSig3");
    var colSig4 = tableSignatures.addColumn("colSig4");
    var colSig5 = tableSignatures.addColumn("colSig5");
    var colSig6 = tableSignatures.addColumn("colSig6");

    //Select from the journal all the given rows, take all the needed data, calculate totals and create the vouchers print
    for (j = 0; j < rowsToProcess.length; j++) 
    {
        for (i = 0; i < journal.rowCount; i++)
        {
            if (i == rowsToProcess[j])
            {
                var tRow = journal.row(i);
                date = tRow.value("Date");
                doc = tRow.value("Doc");
                desc = tRow.value("Description");
                var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
                
                // Debit
                if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 )
                {
                    totDebit = Banana.SDecimal.add(totDebit, amount, {'decimals':2});
                }
                
                // Credit
                else {
                    totCredit = Banana.SDecimal.add(totCredit, amount, {'decimals':2});
                }
                //report.addParagraph("Array: " + rowsToProcess[j] + ", Journal: " + i + ", " + date + ", " + desc + ", " + amount + ", " + totDebit + ", " + totCredit);
            }
        }
    }

    /* 1. Print the date and the voucher number */
    printInfoVoucher(tableInfo, report, docNumber, date);

    /* 2. Print the all the transactions data */
    printTransactions(table, journal, line, rowsToProcess);

    /* 3. Print the total line */
    printTotal(table, totDebit, totCredit, report);

    /* 4. Print the signatures line */
    printSignatures(tableSignatures, report);
}


/* Function that prints the date and the voucher number */
function printInfoVoucher(tableInfo, report, docNumber, date) {

    //Title
    tableRow = tableInfo.addRow();
    tableRow.addCell(getValue(param, "voucher", "chinese"), "heading1 alignCenter", 11);

    tableRow = tableInfo.addRow();
    tableRow.addCell(getValue(param, "voucher", "english"), "heading2 alignCenter", 11);


    //Date and Voucher number
    tableRow = tableInfo.addRow();
    
    var cell1 = tableRow.addCell("", "", 1);
    cell1.addParagraph(" ", "");

    var cell2 = tableRow.addCell("", "alignRight border-top-double", 1);
    cell2.addParagraph(getValue(param, "date", "chinese") + ": ");

    var d = Banana.Converter.toDate(date);
    var year = d.getFullYear();
    var month = d.getMonth()+1;
    var day = d.getDate();

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

    tableRow.addCell(getValue(param, "description", "chinese"), "alignCenter border-left-black border-top-black border-right", 1);
    tableRow.addCell(getValue(param, "genLedAc", "chinese"), "alignCenter border-left border-top-black border-right", 1);
    tableRow.addCell(getValue(param, "subLedAc", "chinese"), "alignCenter border-left border-top-black border-right-black", 1);
    tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
    tableRow.addCell(getValue(param, "debitAmount", "chinese"), "alignCenter border-left border-top-black border-right-black", 11);
    tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
    tableRow.addCell(getValue(param, "creditAmount", "chinese"), "alignCenter border-left border-top-black border-right-black",11);
    tableRow.addCell("", "border-top-black border-left-black border-right-black", 1);
    tableRow.addCell(res[0], "alignCenter border-left border-top-black border-right-black", 1);

    tableRow = tableHeader.addRow();
    tableRow.addCell(getValue(param, "description", "english"), "alignCenter border-left-black border-right", 1);
    tableRow.addCell(getValue(param, "genLedAc", "english"), "alignCenter border-left border-right", 1);
    tableRow.addCell(getValue(param, "subLedAc", "english"), "alignCenter border-left border-right-black", 1);
    tableRow.addCell("", "border-left-black border-right-black", 1);
    tableRow.addCell(getValue(param, "debitAmount", "english"), "alignCenter border-left border-right-black border-bottom", 11);
    tableRow.addCell("", "border-left-black border-right-black", 1);
    tableRow.addCell(getValue(param, "creditAmount", "english"), "alignCenter border-left border-right-black border-bottom",11);
    tableRow.addCell("", "border-left-black border-right-black", 1);
    tableRow.addCell(res[1], "alignCenter border-left border-right-black", 1);


    //Row with the digits titles
    tableRow = tableHeader.addRow();

    tableRow.addCell("", "border-left-black border-right border-bottom-black", 1);
    tableRow.addCell("", "border-left border-right border-bottom-black", 1);
    tableRow.addCell("", "border-left border-right-black border-bottom-black", 1);

    tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);

    tableRow.addCell("亿", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("千", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("百", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("十", "font-size-digits alignCenter padding-right border-left-1px border-right border-bottom-black", 1);
    tableRow.addCell("万", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("千", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("百", "font-size-digits alignCenter padding-right border-left-1px border-right border-bottom-black", 1);
    tableRow.addCell("十", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("元", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("角", "font-size-digits alignCenter padding-right border-left-1px border-right border-bottom-black", 1);
    tableRow.addCell("分", "font-size-digits alignCenter padding-right border-left border-right-black border-bottom-black", 1);

    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom-black", 1);

    tableRow.addCell("亿", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("千", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("百", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("十", "font-size-digits alignCenter padding-right border-left-1px border-right border-bottom-black", 1);
    tableRow.addCell("万", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("千", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("百", "font-size-digits alignCenter padding-right border-left-1px border-right border-bottom-black", 1);
    tableRow.addCell("十", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("元", "font-size-digits alignCenter padding-right border-left border-right border-bottom-black", 1);
    tableRow.addCell("角", "font-size-digits alignCenter padding-right border-left-1px border-right border-bottom-black", 1);
    tableRow.addCell("分", "font-size-digits alignCenter padding-right border-left border-right-black border-bottom-black", 1);

    tableRow.addCell("", "border-left-black border-right-black border-bottom-black", 1);

    tableRow.addCell(getValue(param, "pr", "english"), "border-left border-right-black border-bottom-black", 1);


    /* Print transactions rows */
    var tmpDescription;
    for (j = 0; j < rowsToProcess.length; j++) 
    {
        for (i = 0; i < journal.rowCount; i++)
        {
            if (i == rowsToProcess[j])
            {
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
                
                // Debit
                if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ) {

                    getDigits(Banana.Converter.toLocaleNumberFormat(amount), tableRow, false);
                    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-black border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right-black border-bottom", 1);
                }
                // Credit
                else {

                    tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-black border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
                    tableRow.addCell("", "text-black padding-right border-left border-top border-right-black border-bottom", 1);
                    getDigits(Banana.Converter.toLocaleNumberFormat(amount), tableRow, false);
                }

                tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
                tableRow.addCell(" ", "text-black alignCenter border-left border-top border-right-black border-bottom", 1);
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
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right-black border-bottom", 1);

        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell("*", "text-black padding-right border-left border-top border-right-black border-bottom", 1);

        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
        tableRow.addCell(" ", "text-black alignCenter border-left border-top border-right-black border-bottom", 1);
        
        line++;
    }
}


/* Function that prints the total line of the voucher */
function printTotal(table, totDebit, totCredit, report) {

    //Attachments
    tableRow = table.addRow();
    tableRow.addCell(getValue(param, "attachments", "chinese") + "   " + getValue(param, "attachments", "english"), "alignCenter padding-left border-left-black border-top-black border-right border-bottom-black", 1);

    //Total
    tableRow.addCell(getValue(param, "total", "chinese") + "   " + getValue(param, "total", "english"), "alignCenter padding-left border-left border-top-black border-right-black border-bottom-black", 2);
    getDigits(Banana.Converter.toLocaleNumberFormat(totDebit), tableRow, true);
    getDigits(Banana.Converter.toLocaleNumberFormat(totCredit), tableRow, true);
    tableRow.addCell("", "border-left-black border-top-black border-right-black border-bottom-black", 1);
    tableRow.addCell(" ", "text-black alignCenter border-left border-top-black border-right-black border-bottom-black", 1);
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


/* Function that takes a number and save all the digits */
function getDigits(num, table, isTotalLine) {

    //Create all the digits
    var a8 = "";
    var a7 = "";
    var a6 = "";
    var a5 = "";
    var a4 = "";
    var a3 = "";
    var a2 = "";
    var a1 = "";
    var a0 = "";
    var c1 = "";
    var c2 = "";

    //replace everything except numbers
    var arr = num.replace(/\D/g,'');

    var output = [];
    for (var i = 0, len = num.length; i < len; i += 1) {
        output.push(+num.charAt(i));
    }
    //example number 1'250.37 =>  output=[1,2,5,0,3,7]
    
    //Save each element of the array in the corresponding variable
    for (var i = 0; i < arr.length; i++) {
        
        ////
        if (arr[arr.length - 1]) {    
            c2 = arr[arr.length - 1];
        }

        if (arr[arr.length - 2]) {
            c1 = arr[arr.length - 2];
        }

        ////
        if (arr[arr.length - 3]) {
            a0 = arr[arr.length - 3];
        }

        if (arr[arr.length - 4]) {
            a1 = arr[arr.length - 4];
        } else {
            a1 = "¥";
        }
            
        if (arr[arr.length - 5]) {
            a2 = arr[arr.length - 5];
        } else {
            if (a1 !== "¥") {
                a2 = "¥";
            }
        }
        
        ////
        if (arr[arr.length - 6]) {
            a3 = arr[arr.length - 6];
        } else {
            if (a1 !== "¥" && a2 !== "¥") {
                a3 = "¥";
            }
        }
            
        if (arr[arr.length - 7]) {
            a4 = arr[arr.length - 7];
        } else {
            if (a1 !== "¥" && a2 !== "¥" && a3 !== "¥") {
                a4 = "¥";
            }
        }
            
        if (arr[arr.length - 8]) {
            a5 = arr[arr.length - 8];
        } else {
            if (a1 !== "¥" && a2 !== "¥" && a3 !== "¥" && a4 !== "¥") {
                a5 = "¥";
            }
        }

        ////
        if (arr[arr.length - 9]) {
            a6 = arr[arr.length - 9];
        } else {
            if (a1 !== "¥" && a2 !== "¥" && a3 !== "¥" && a4 !== "¥" && a5 !== "¥") {
                a6 = "¥";
            }
        }
            
        if (arr[arr.length - 10]) {
            a7 = arr[arr.length - 10];
        } else {
            if (a1 !== "¥" && a2 !== "¥" && a3 !== "¥" && a4 !== "¥" && a5 !== "¥" && a6 !== "¥") {
                a7 = "¥";
            }
        }
           
        if (arr[arr.length - 11]) {
            a8 = arr[arr.length - 11];
        } else {
            if (a1 !== "¥" && a2 !== "¥" && a3 !== "¥" && a4 !== "¥" && a5 !== "¥" && a6 !== "¥" && a7 !== "¥") {
                a8 = "¥";
            }
        }
    }

    if (!isTotalLine) {
        tableRow.addCell("", "border-top border-left-black border-right-black border-bottom", 1);
        tableRow.addCell(a8, "text-black padding-right alignRight border-left-black border-top border-right border-bottom", 1);
        tableRow.addCell(a7, "text-black padding-right alignRight border-left border-top border-right border-bottom", 1);
        tableRow.addCell(a6, "text-black padding-right alignRight border-left border-top border-right border-bottom", 1);
        tableRow.addCell(a5, "text-black padding-right alignRight border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell(a4, "text-black padding-right alignRight border-left border-top border-right border-bottom", 1);
        tableRow.addCell(a3, "text-black padding-right alignRight border-left border-top border-right border-bottom", 1);
        tableRow.addCell(a2, "text-black padding-right alignRight border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell(a1, "text-black padding-right alignRight border-left border-top border-right border-bottom", 1);
        tableRow.addCell(a0, "text-black padding-right alignRight border-left border-top border-right border-bottom", 1);
        tableRow.addCell(c1, "text-black padding-right alignRight border-left-1px border-top border-right border-bottom", 1);
        tableRow.addCell(c2, "text-black padding-right alignRight border-left border-top border-right-black border-bottom", 1);
    }
    else {
        tableRow.addCell("", "border-top-black border-left-black border-right-black border-bottom-black", 1);
        tableRow.addCell(a8, "text-black padding-right alignRight border-left-black border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a7, "text-black padding-right alignRight border-left border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a6, "text-black padding-right alignRight border-left border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a5, "text-black padding-right alignRight border-left-1px border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a4, "text-black padding-right alignRight border-left border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a3, "text-black padding-right alignRight border-left border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a2, "text-black padding-right alignRight border-left-1px border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a1, "text-black padding-right alignRight border-left border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(a0, "text-black padding-right alignRight border-left border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(c1, "text-black padding-right alignRight border-left-1px border-top-black border-right border-bottom-black", 1);
        tableRow.addCell(c2, "text-black padding-right alignRight border-left border-top-black border-right-black border-bottom-black", 1);
    }
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
    stylesheet.addStyle("body", "font-family : Times New Roman; font-size:10pt; color:" + generalParam.textColor);
    stylesheet.addStyle(".font-size-digits", "font-size:6pt");
    stylesheet.addStyle(".text-green", "color:" + generalParam.textColor);
    stylesheet.addStyle(".text-black", "color:black");

    stylesheet.addStyle(".border-top", "border-top: thin solid " +  generalParam.textColor);
    stylesheet.addStyle(".border-right", "border-right: thin solid " +  generalParam.textColor);
    stylesheet.addStyle(".border-bottom", "border-bottom: thin solid " +  generalParam.textColor);
    stylesheet.addStyle(".border-left", "border-left: thin solid " +  generalParam.textColor);
    stylesheet.addStyle(".border-left-1px", "border-left: 1px solid " +  generalParam.textColor);
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
    stylesheet.addStyle(".c1", "width:40%");
    stylesheet.addStyle(".c2", "width:12%");
    stylesheet.addStyle(".c3", "width:12%");
    stylesheet.addStyle(".cs1", "width:0.4%");
    stylesheet.addStyle(".c4", "width:1.5%");
    stylesheet.addStyle(".c5", "width:1.5%");
    stylesheet.addStyle(".c6", "width:1.5%");
    stylesheet.addStyle(".c7", "width:1.5%");
    stylesheet.addStyle(".c8", "width:1.5%");
    stylesheet.addStyle(".c9", "width:1.5%");
    stylesheet.addStyle(".c10", "width:1.5%");
    stylesheet.addStyle(".c11", "width:1.5%");
    stylesheet.addStyle(".c12", "width:1.5%");
    stylesheet.addStyle(".c13", "width:1.5%");
    stylesheet.addStyle(".c14", "width:1.5%");
    stylesheet.addStyle(".cs2", "width:0.4%");
    stylesheet.addStyle(".c15", "width:1.5%");
    stylesheet.addStyle(".c16", "width:1.5%");
    stylesheet.addStyle(".c17", "width:1.5%");
    stylesheet.addStyle(".c18", "width:1.5%");
    stylesheet.addStyle(".c19", "width:1.5%");
    stylesheet.addStyle(".c20", "width:1.5%");
    stylesheet.addStyle(".c21", "width:1.5%");
    stylesheet.addStyle(".c22", "width:1.5%");
    stylesheet.addStyle(".c23", "width:1.5%");
    stylesheet.addStyle(".c24", "width:1.5%");
    stylesheet.addStyle(".c25", "width:1.5%");
    stylesheet.addStyle(".cs3", "width:0.4%");
    stylesheet.addStyle(".c26", "width:3%");

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


