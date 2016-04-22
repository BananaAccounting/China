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
// @id = ch.banana.addon.reportvoucherchina2016
// @api = 1.0
// @pubdate = 2016-04-022
// @publisher = Banana.ch SA
// @description = Report Voucher China
// @description.cn = Report Voucher China
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @docproperties = 
// @outputformat = none
// @inputdataform = none
// @timeout = -1

var param = {};
param = {

    /*********
        GREEN
    *********/
    //internal             //Chinese
    "insertDocument"    : "Insert a document number",
    "transferVoucher"   : "Transfer Voucher",
    "date"              : "Date",
    "voucherNumber"     : "Voucher Number",
    "explanation"       : "Explanation",
    "account"           : "Account",
    "amount"            : "Amount",
    "debit"             : "Debit",
    "credit"            : "Credit",
    "total"             : "Total",
    "approved"          : "Approved",
    "checked"           : "Checked",
    "entered"           : "Entered",
    "prepared"          : "Prepared",

    /*******
        RED
    *******/
    //internal             //Chinese
    "debitVoucher"      : "Debit Voucher",
    "debitAccount"      : "Debit Account",
    //Solo colonna debit





    /********
        BLUE
    ********/
    //internal             //Chinese
    "creditVoucher"      : "Credit Voucher",
    "creditAccount"      : "Credit Account"
    //Solo colonna credit


}



var colorParam = {
    "textColor" : ""
}




function exec(string) {

    //Create the report
    var report = Banana.Report.newReport(param.transferVoucher);
    
    try {

        //Table Transactions
        var transactions = Banana.document.table('Transactions');
        var itemSelected = Banana.Ui.getItem('Input', 'Choose a report', [param.transferVoucher,'Debit Voucher','Credit Voucher'], 0, false);
        var docNumber = Banana.Ui.getText(param.insertDocument, param.insertDocument,'');

        if (itemSelected === param.transferVoucher) {
            colorParam.textColor = "green";
            createTransferVoucherReport(report, docNumber);
        }
        else if (itemSelected === param.debitVoucher) {
            colorParam.textColor = "red";
            createDebitVoucherReport(report, docNumber);
        }
        else if (itemSelected === param.creditVoucher) {
            colorParam.textColor = "blue";
            createCreditVoucherReport(report, docNumber);
        }



    } catch (e) {
        //If the selected row doesn't contains the Doc value it is displayed a dialog window  
        Banana.Ui.showInformation('Doc not found', 'Doc not found. Please select a row with values from the Transactions table.');
        return;
    }

    //Add the styles
    var stylesheet = createStyleSheet();
    Banana.Report.preview(report, stylesheet);
}


/*
    1. Transfer Voucher
*/
function createTransferVoucherReport(report, docNumber) {
    
    addFooter(report);

    //Table Journal
    var journal = Banana.document.journal(Banana.document.ORIGINTYPE_CURRENT, Banana.document.ACCOUNTTYPE_NORMAL);
 
    var date = "";
    var doc = ""; 
    var desc = "";
    var totDebit = "";
    var totCredit = "";

    for (i = 0; i < journal.rowCount; i++)
    {
        var tRow = journal.row(i);

        //Take the row from the journal that has the same "Doc" as the one selectet by the user (cursor)
        if (docNumber && tRow.value('Doc') === docNumber)
        {
            date = tRow.value("Date");
            doc = tRow.value("Doc");
            desc = tRow.value("Description");

            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            // Debit
            if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ){
                totDebit = Banana.SDecimal.add(totDebit, amount, {'decimals':0});
            }
            // Credit
            else {
                totCredit = Banana.SDecimal.add(totCredit, amount, {'decimals':0});
            }
        }
    }

    report.addParagraph(param.transferVoucher, "heading1 alignCenter");
    report.addParagraph(" ", "");
    report.addParagraph(" ", "");


    /**
        TABLE 1
    **/
    var table = report.addTable("table");

    tableRow = table.addRow();
    tableRow.addCell("                                                         ", "", 1);

    tableRow.addCell(param.date + ": ", "alignCenter", 1);
    tableRow.addCell(Banana.Converter.toLocaleDateFormat(date), "", 1);

    tableRow.addCell("                                                         ", "", 1);

    tableRow.addCell(param.voucherNumber + ": ", "padding-left border-left border-top border-right border-bottom", 1);
    tableRow.addCell(docNumber, "padding-left border-top border-right border-bottom", 1);

    report.addParagraph(" ", "");
    report.addParagraph(" ", "");

 

    /**
        TABLE 2
    **/
    var table = report.addTable("table");

    tableRow = table.addRow();
    tableRow.addCell(param.explanation, " alignCenter border-left border-top border-right ", 1);
    tableRow.addCell(param.account, " alignCenter border-left border-top border-right", 1);
    tableRow.addCell(param.amount, " alignCenter border-left border-top border-right border-bottom", 2);

    tableRow = table.addRow();
    tableRow.addCell("", "border-left border-right border-bottom", 1);
    tableRow.addCell("", "border-left border-right border-bottom", 1);
    tableRow.addCell(param.debit, " alignCenter border-left border-top border-right border-bottom", 1);
    tableRow.addCell(param.credit, " alignCenter border-left border-top border-right border-bottom", 1);
    
    //Print transactions row
    for (i = 0; i < journal.rowCount; i++)
    {
        var tRow = journal.row(i);

        //Take the row from the journal that has the same "Doc" as the one selectet by the user (cursor)
        if (tRow.value('Doc') === docNumber)
        {
            tableRow = table.addRow();  
            tableRow.addCell(tRow.value('Description'), "padding-left border-left border-top border-right border-bottom", 1);
            tableRow.addCell(tRow.value('JAccount'), "padding-left border-left border-top border-right border-bottom", 1);
            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            // Debit
            if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ){
                tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amount), "padding-right alignRight  border-left border-top border-right border-bottom", 1);
                tableRow.addCell("", "padding-right border-left border-top border-right border-bottom", 1);
            }
            // Credit
            else {
                tableRow.addCell("", "padding-right border-left border-top border-right border-bottom", 1);
                tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amount), "padding-right alignRight border-left border-top border-right border-bottom", 1);
            }

        }
    }

    //Total Line
    tableRow = table.addRow();
    tableRow.addCell(param.total,"padding-left bold border-left border-top border-right border-bottom", 2);
    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(totDebit), "padding-right bold alignRight border-left border-top border-right border-bottom", 1);
    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(totCredit), "padding-right bold alignRight border-left border-top border-right border-bottom", 1);

    report.addParagraph(" ", "");


    /**
        TABLE 7
    **/
    var table = report.addTable("table");
    tableRow = table.addRow();
    tableRow.addCell(param.approved + ": ", "", 1);
    tableRow.addCell(param.checked + ": ", "", 1);
    tableRow.addCell(param.entered + ": ", "", 1);
    tableRow.addCell(param.prepared + ": ", "", 1);
}





/* 
    2. Debit Voucher
*/
function createDebitVoucherReport(report, docNumber) {
    
    addFooter(report);

    //Table Journal
    var journal = Banana.document.journal(Banana.document.ORIGINTYPE_CURRENT, Banana.document.ACCOUNTTYPE_NORMAL);
 
    var date = "";
    var doc = ""; 
    var desc = "";
    var totDebit = "";
    var totCredit = "";

    for (i = 0; i < journal.rowCount; i++)
    {
        var tRow = journal.row(i);

        //Take the row from the journal that has the same "Doc" as the one selectet by the user (cursor)
        if (docNumber && tRow.value('Doc') === docNumber)
        {
            date = tRow.value("Date");
            doc = tRow.value("Doc");
            desc = tRow.value("Description");

            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            // Debit
            if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ){
                totDebit = Banana.SDecimal.add(totDebit, amount, {'decimals':0});
            }
            // Credit
            else {
                totCredit = Banana.SDecimal.add(totCredit, amount, {'decimals':0});
            }
        }
    }

    report.addParagraph(param.debitVoucher, "heading1 alignCenter");
    report.addParagraph(" ", "");
    report.addParagraph(" ", "");


    /**
        TABLE 1
    **/
    var table = report.addTable("table");

    tableRow = table.addRow();
    tableRow.addCell(param.debitAccount + ": ", "", 1);
    tableRow.addCell("xxxxx", "", 1);

    tableRow.addCell(param.date + ": ", "alignCenter", 1);
    tableRow.addCell(Banana.Converter.toLocaleDateFormat(date), "", 1);

    tableRow.addCell("                                                         ", "", 1);

    tableRow.addCell(param.voucherNumber + ": ", "padding-left border-left border-top border-right border-bottom", 1);
    tableRow.addCell(docNumber, "padding-left border-top border-right border-bottom", 1);

    report.addParagraph(" ", "");
    report.addParagraph(" ", "");

 

    /**
        TABLE 2
    **/
    var table = report.addTable("table");

    tableRow = table.addRow();
    tableRow.addCell(param.explanation, " alignCenter border-left border-top border-right ", 1);
    tableRow.addCell(param.account, " alignCenter border-left border-top border-right", 1);
    tableRow.addCell("??", " alignCenter border-left border-top border-right border-bottom", 1);
    tableRow.addCell(param.debit, " alignCenter border-left border-top border-right border-bottom", 1);

    
    //Print transactions row
    for (i = 0; i < journal.rowCount; i++)
    {
        var tRow = journal.row(i);

        //Take the row from the journal that has the same "Doc" as the one selectet by the user (cursor)
        if (tRow.value('Doc') === docNumber && Banana.SDecimal.sign(tRow.value('JAmount')) > 0 )
        {
            tableRow = table.addRow();  
            tableRow.addCell(tRow.value('Description'), "padding-left border-left border-top border-right border-bottom", 1);
            tableRow.addCell(tRow.value('JAccount'), "padding-left border-left border-top border-right border-bottom", 1);
            tableRow.addCell("??", "padding-left border-left border-top border-right border-bottom", 1);
            
            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amount), "padding-right alignRight border-left border-top border-right border-bottom", 1);
            
        }
    }

    //Total Line
    tableRow = table.addRow();
    tableRow.addCell(param.total,"padding-left bold border-left border-top border-right border-bottom", 3);
    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(totDebit), "padding-right bold alignRight border-left border-top border-right border-bottom", 1);

    report.addParagraph(" ", "");


    /**
        TABLE 7
    **/
    var table = report.addTable("table");
    tableRow = table.addRow();
    tableRow.addCell(param.approved + ": ", "", 1);
    tableRow.addCell(param.checked + ": ", "", 1);
    tableRow.addCell(param.entered + ": ", "", 1);
    tableRow.addCell(param.prepared + ": ", "", 1);
}

/*
    3. Credit Voucher
*/
function createCreditVoucherReport(report, docNumber) {
    
    addFooter(report);

    //Table Journal
    var journal = Banana.document.journal(Banana.document.ORIGINTYPE_CURRENT, Banana.document.ACCOUNTTYPE_NORMAL);
 
    var date = "";
    var doc = ""; 
    var desc = "";
    var totDebit = "";
    var totCredit = "";

    for (i = 0; i < journal.rowCount; i++)
    {
        var tRow = journal.row(i);

        //Take the row from the journal that has the same "Doc" as the one selectet by the user (cursor)
        if (docNumber && tRow.value('Doc') === docNumber)
        {
            date = tRow.value("Date");
            doc = tRow.value("Doc");
            desc = tRow.value("Description");

            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            // Debit
            if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ){
                totDebit = Banana.SDecimal.add(totDebit, amount, {'decimals':0});
            }
            // Credit
            else {
                totCredit = Banana.SDecimal.add(totCredit, amount, {'decimals':0});
            }
        }
    }

    report.addParagraph(param.creditVoucher, "heading1 alignCenter");
    report.addParagraph(" ", "");
    report.addParagraph(" ", "");


    /**
        TABLE 1
    **/
    var table = report.addTable("table");

    tableRow = table.addRow();
    tableRow.addCell(param.creditAccount + ": ", "", 1);
    tableRow.addCell("xxxxx", "", 1);

    tableRow.addCell(param.date + ": ", "alignCenter", 1);
    tableRow.addCell(Banana.Converter.toLocaleDateFormat(date), "", 1);

    tableRow.addCell("                                                         ", "", 1);

    tableRow.addCell(param.voucherNumber + ": ", "padding-left border-left border-top border-right border-bottom", 1);
    tableRow.addCell(docNumber, "padding-left border-top border-right border-bottom", 1);

    report.addParagraph(" ", "");
    report.addParagraph(" ", "");

 

    /**
        TABLE 2
    **/
    var table = report.addTable("table");

    tableRow = table.addRow();
    tableRow.addCell(param.explanation, " alignCenter border-left border-top border-right ", 1);
    tableRow.addCell(param.account, " alignCenter border-left border-top border-right", 1);
    tableRow.addCell("??", " alignCenter border-left border-top border-right border-bottom", 1);
    tableRow.addCell(param.credit, " alignCenter border-left border-top border-right border-bottom", 1);

    
    //Print transactions row
    for (i = 0; i < journal.rowCount; i++)
    {
        var tRow = journal.row(i);

        //Take the row from the journal that has the same "Doc" as the one selectet by the user (cursor)
        if (tRow.value('Doc') === docNumber && Banana.SDecimal.sign(tRow.value('JAmount')) <= 0 )
        {
            tableRow = table.addRow();  
            tableRow.addCell(tRow.value('Description'), "padding-left border-left border-top border-right border-bottom", 1);
            tableRow.addCell(tRow.value('JAccount'), "padding-left border-left border-top border-right border-bottom", 1);
            tableRow.addCell("??", "padding-left border-left border-top border-right border-bottom", 1);
            
            var amount = Banana.SDecimal.abs(tRow.value('JAmount'));
            tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amount), "padding-right alignRight border-left border-top border-right border-bottom", 1);
            
        }
    }

    //Total Line
    tableRow = table.addRow();
    tableRow.addCell(param.total,"padding-left bold border-left border-top border-right border-bottom", 3);
    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(totCredit), "padding-right bold alignRight border-left border-top border-right border-bottom", 1);

    report.addParagraph(" ", "");


    /**
        TABLE 7
    **/
    var table = report.addTable("table");
    tableRow = table.addRow();
    tableRow.addCell(param.approved + ": ", "", 1);
    tableRow.addCell(param.checked + ": ", "", 1);
    tableRow.addCell(param.entered + ": ", "", 1);
    tableRow.addCell(param.prepared + ": ", "", 1);
}




















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


//This function adds a Footer to the report
function addFooter(report) {
   report.getFooter().addClass("footer");
   var versionLine = report.getFooter().addText("Banana 财务会计软件8", "description");


}

//The main purpose of this function is to create styles for the report print
function createStyleSheet() {
    //Create stylesheet
    var stylesheet = Banana.Report.newStyleSheet();
    
    //Set page layout
    var pageStyle = stylesheet.addStyle("@page");

    //Set the margins
    pageStyle.setAttribute("margin", "15mm 15mm 10mm 20mm");
    //pageStyle.setAttribute("size", "landscape");

    stylesheet.addStyle("body", "font-family : Helvetica; font-size:12pt; color:" + colorParam.textColor);
    stylesheet.addStyle(".border-top", "border-top: thin solid " +  colorParam.textColor);
    stylesheet.addStyle(".border-right", "border-right: thin solid " +  colorParam.textColor);
    stylesheet.addStyle(".border-bottom", "border-bottom: thin solid " +  colorParam.textColor);
    stylesheet.addStyle(".border-left", "border-left: thin solid " +  colorParam.textColor);
    stylesheet.addStyle(".padding-left", "padding-left:5px");
    stylesheet.addStyle(".padding-right", "padding-right:5px");

    style = stylesheet.addStyle(".heading1");
    style.setAttribute("font-size", "14px");
    style.setAttribute("font-weight", "bold");
    
    //Set Table styles
    style = stylesheet.addStyle("table");
    style.setAttribute("width", "100%");
    style.setAttribute("font-size", "8px");
    stylesheet.addStyle("table.table td", "padding-bottom: 2px; padding-top: 5px");
    //stylesheet.addStyle("table.table td", "border: thin solid colorParam.textColor; padding-bottom: 2px; padding-top: 5px");

    style = stylesheet.addStyle(".bold");
    style.setAttribute("font-weight", "bold");

    style = stylesheet.addStyle(".alignRight");
    style.setAttribute("text-align", "right");

    style = stylesheet.addStyle(".alignCenter");
    style.setAttribute("text-align", "center");

    style = stylesheet.addStyle(".footer");
    style.setAttribute("text-align", "right");
    style.setAttribute("font-size", "8px");
    style.setAttribute("font-family", "Courier New");

    
    return stylesheet;
}


