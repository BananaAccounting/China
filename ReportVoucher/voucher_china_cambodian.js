// Copyright [2020] [Banana.ch SA - Lugano Switzerland]
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
// @id = ch.banana.addon.voucherchinas.cambodian
// @api = 1.0
// @pubdate = 2020-10-14
// @publisher = Banana.ch SA
// @description.zh = 柬埔寨会计凭证 (中文-柬埔寨文)
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @docproperties = 
// @outputformat = none
// @inputdataform = none
// @timeout = -1


//-----------------------------------------------------------
//  General parameters that can be changed by the user:
//    - textColor: the color of text and table borders. Default value is "#32CD32"
//    - numberOfRows: the number of transactions rows that are printed on each voucher. Default value is 6
//    - pageSize: the size of the page (e.g. "A5 landscape"). Default value is "A4"
//-----------------------------------------------------------
var generalParam = {};
function setGeneralParam(userParam) {

    generalParam.numberOfRows = 5;
    generalParam.rowHeight = "13mm";
    generalParam.rowHeightTotal = "9mm";
    generalParam.column1 = "64mm";
    generalParam.column2 = "60mm";
    generalParam.column3 = "35mm";
    generalParam.column4 = "35mm";
    generalParam.fontSize = "8px";
    generalParam.fontSizeTable = "8px";

    // Color    
    if (!userParam.color) {
        userParam.color = "Black";
    }
    generalParam.textColor = userParam.color;

    // Page size
    if (userParam.custompagesize) {
        //remove all non-digits
        userParam.customwidth = userParam.customwidth.replace(/[^0-9]/g,'');
        userParam.customheight = userParam.customheight.replace(/[^0-9]/g,'');

        var landscape = "";
        if (userParam.landscape) {
            landscape = " landscape";
        }

        generalParam.pageSize = userParam.customheight + "mm " + userParam.customwidth + "mm" + landscape;
    }
    else {
        if (userParam.pagesize0 && !userParam.pagesize1 && !userParam.pagesize2 && !userParam.pagesize3 && !userParam.pagesize4) {
            generalParam.pageSize = "140mm 240mm";//"84mm 200mm landscape";
        }
        else if (userParam.pagesize1 && !userParam.pagesize0 && !userParam.pagesize2 && !userParam.pagesize3 && !userParam.pagesize4) {
            generalParam.pageSize = "142mm 243mm landscape";
        }
        else if (userParam.pagesize2 && !userParam.pagesize0 && !userParam.pagesize1 && !userParam.pagesize3 && ! userParam.pagesize4) {
            generalParam.pageSize = "127mm 210mm landscape";
        }
        else if (userParam.pagesize3 && !userParam.pagesize0 && !userParam.pagesize1 && !userParam.pagesize2 && !userParam.pagesize4) {
            generalParam.pageSize = "A5 landscape";
        }
        else {
            generalParam.pageSize = "A4";
        }
    }
}


/* Function that loads the texts parameters */
var param = [];
function loadParam() {
    param.push({"id": "showInformation", "cambodia":"Document number not found. Please insert a valid value from the Document column of the Transactions table", "chinese":"未找到文件编号，请您从发生业务表格的文件列中选择一个有效值。"});
    param.push({"id": "printVoucher", "cambodia":"Print Voucher", "chinese":"打印凭证"});
    param.push({"id": "insertDocument", "cambodia":"Insert a document number (or insert * to print all the vouchers)", "chinese":"输入文件编号，或输入'*'号用来打印所有的凭证。"});
    param.push({"id": "voucher", "cambodia":"កំណត់ត្រាបញ្ជីទូទាត់", "chinese":"记   账   凭   证"});
    param.push({"id": "date", "cambodia":"កាលបរិច្ឆេទ", "chinese":"日期"});
    param.push({"id": "year", "cambodia":"ឆ្នាំ", "chinese":"年"});
    param.push({"id": "month", "cambodia":"ខែ", "chinese":"月"});
    param.push({"id": "day", "cambodia":"ថ្ងៃទី", "chinese":"日"});
    param.push({"id": "voucherNumber", "cambodia":"លេរៀង", "chinese":"第   号"});
    param.push({"id": "description", "cambodia":"បរិយាយ", "chinese":"摘要"});
    param.push({"id": "genLedAc", "cambodia":"គណន", "chinese":"科目"});
    param.push({"id": "debitAmount", "cambodia":"ឥណពន្ធ", "chinese":"借方金额"});
    param.push({"id": "creditAmount", "cambodia":"ឥណទាន", "chinese":"贷方金额"});
    param.push({"id": "attachments", "cambodia":"ឧបសម្ព័ន្ធ សន្លឹក", "chinese":"附单据     张"});
    param.push({"id": "total", "cambodia":"ចំនួនសរុប", "chinese":"合计"});
    param.push({"id": "approved", "cambodia":"អ្នកអនុម័ត", "chinese":"核      准"});
    param.push({"id": "checked", "cambodia":"អ្នកត្រួតពិនិត្យ", "chinese":"复      核"});
    param.push({"id": "entered", "cambodia":"អ្នករៀបបញ្ជី", "chinese":"记      账"});
    param.push({"id": "cashier", "cambodia":"បេឡាករ", "chinese":"出      纳"});
    param.push({"id": "prepared", "cambodia":"អ្នកចុះបញ្ជី", "chinese":"制      单"});
    param.push({"id": "receiver", "cambodia":"អ្នកទទួល", "chinese":"签      收"});
}

/* Main function */
function exec(string) {

    if (!Banana.document) {
        return "@Cancel";
    }
    
    // Load all the parameters
    loadParam();

    // Settings
    var userParam = initUserParam();

    // Retrieve saved param
    var savedParam = Banana.document.getScriptSettings();
    if (savedParam && savedParam.length > 0) {
        userParam = JSON.parse(savedParam);
    }
    // If needed show the settings dialog to the user
    if (!options || !options.useLastSettings) {
        userParam = settingsDialog(); // From properties
    }
    if (!userParam) {
        return "@Cancel";
    }
    
    // Set General parameters
    setGeneralParam(userParam);

    var report = printVoucher(Banana.document, userParam);

    //Add the styles
    var stylesheet = createStyleSheet();
    Banana.Report.preview(report, stylesheet);
}

/* Function that creates an array of rows and use it to print the voucher report */
function printVoucher(banDoc, userParam) {

    //Create the report
    var report = Banana.Report.newReport(param.transferVoucher);

    //Create the Journal table which contains all the data of the accounting
    var journal = banDoc.journal(banDoc.ORIGINTYPE_CURRENT, banDoc.ACCOUNTTYPE_NORMAL);

    //Create a list with all the doc numbers of the Transaction table
    var listDoc = getDocList(banDoc);
    
    if (userParam.voucher && userParam.voucher !== "*") {
        //Check if the userParam.voucher exists in the listDoc array
        if (listDoc.indexOf(userParam.voucher) > -1) {
            listDoc = [];
            listDoc.push(userParam.voucher);
        }
        else { //userParam.voucher doesn't exists
            listDoc = [];
            Banana.Ui.showInformation("", getValue(param, "showInformation", "chinese"));
            return "@Cancel";
        }
    }
    else if (!userParam.voucher) {
        listDoc = [];
        return "@Cancel";
    }

    if (listDoc.length > 0) {
        //Array to store all the rows that will be processed
        var rowsToProcess = []; 
        for (var j = 0; j < listDoc.length; j++) {
            var docNumber = listDoc[j];

            for (var i = 0; i < journal.rowCount; i++) {
                var tRow = journal.row(i);

                //Check if the actual Doc number of the list is equal to the Doc value of the journal's row
                if (docNumber && tRow.value('Doc') === docNumber) {
                    //Add the row to the array
                    rowsToProcess.push(i);

                    //Check if there are six elements (six rows) into the array.
                    //We want a maximum of six rows per voucher
                    if (rowsToProcess.length == generalParam.numberOfRows) {
                        
                        //Function call to create the voucher for the given rows
                        createVoucherReport(journal, report, docNumber, rowsToProcess, userParam);

                        //Array reset
                        rowsToProcess = [];
                    }
                }
            }
        
            //Check if the array contains something
            if (rowsToProcess.length > 0) {
                //Function call to create the voucher for the given rows
                createVoucherReport(journal, report, docNumber, rowsToProcess, userParam);

                //Array reset
                rowsToProcess = [];
            }
        }
    }

    return report;
}

/* Function that creates the voucher for the given Doc number */
function createVoucherReport(journal, report, docNumber, rowsToProcess, userParam) {

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
    var c4 = table.addColumn("c4");

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
    printInfoVoucher(tableInfo, report, docNumber, date, userParam);

    /* 2. Print the all the transactions data */
    printTransactions(table, journal, line, rowsToProcess, userParam);

    /* 3. Print the total line */
    printTotal(table, totDebit, totCredit, report, userParam);

    /* 4. Print the signatures line */
    printSignatures(tableSignatures, report, userParam);
}

/* Function that prints the date and the voucher number */
function printInfoVoucher(tableInfo, report, docNumber, date, userParam) {

    //Title
    tableRow = tableInfo.addRow();
    var cellTitle = tableRow.addCell("","",11);
    cellTitle.addParagraph(getValue(param, "voucher", "chinese"), "heading1 alignCenter");
    if (userParam.printcambodia) {
        cellTitle.addParagraph(getValue(param, "voucher", "cambodia"), "heading2 alignCenter");
    } else {
      cellTitle.addParagraph("","");  
    }

    //Date and Voucher number
    tableRow = tableInfo.addRow();
    
    var cell1 = tableRow.addCell("", "", 1);
    cell1.addParagraph(" ", "");

    var cell2 = tableRow.addCell("", "heading3 alignRight border-top-double", 1);
    cell2.addParagraph(getValue(param, "date", "chinese") + ": ");
    if (userParam.printcambodia) {
        cell2.addParagraph(getValue(param, "date", "cambodia") + ": ");
    }

    var d = Banana.Converter.toDate(date);
    var year = d.getFullYear();
    var month = d.getMonth()+1;
    var day = d.getDate();

    var cell3 = tableRow.addCell("", "heading3 alignCenter border-top-double", 1);
    cell3.addParagraph(year, "text-black");

    var cell4 = tableRow.addCell("", "heading3 alignCenter border-top-double", 1);
    cell4.addParagraph(getValue(param, "year", "chinese"), "");
    if (userParam.printcambodia) {
        cell4.addParagraph(getValue(param, "year", "cambodia"), "");
    }
    
    var cell5 = tableRow.addCell("", "heading3 alignCenter border-top-double", 1);
    cell5.addParagraph(month, "text-black");

    var cell6 = tableRow.addCell("", "heading3 alignCenter border-top-double", 1);
    cell6.addParagraph(getValue(param, "month", "chinese"), "");
    if (userParam.printcambodia) {
        cell6.addParagraph(getValue(param, "month", "cambodia"), "");
    }

    var cell7 = tableRow.addCell("", "heading3 alignCenter border-top-double", 1);
    cell7.addParagraph(day, "text-black");

    var cell8 = tableRow.addCell("", "heading3 alignCenter border-top-double", 1);
    cell8.addParagraph(getValue(param, "day", "chinese"), "");
    if (userParam.printcambodia) {
        cell8.addParagraph(getValue(param, "day", "cambodia"), "");
    }

    var cell9 = tableRow.addCell("", "", 1);    
    cell9.addParagraph(" ", "");

    var cell10 = tableRow.addCell("", "heading3 padding-left alignCenter", 1);
    cell10.addParagraph(getValue(param, "voucherNumber", "chinese"), "");
    if (userParam.printcambodia) {
        cell10.addParagraph(getValue(param, "voucherNumber", "cambodia"), "");
    }
    
    var cell11 = tableRow.addCell("", "heading3 padding-left", 1);
    if (userParam.printcambodia) {
        cell11.addParagraph(" ", "");
        cell11.addParagraph(docNumber, "text-black");
    } else {
        cell11.addParagraph(docNumber, "text-black");
    }
}

/* Function that prints all the transactions data */
function printTransactions(table, journal, line, rowsToProcess, userParam) {

    /* Table header */
    var tableHeader = table.getHeader();
    tableRow = tableHeader.addRow();

    var cellDescription = tableRow.addCell("","alignCenter heightTrans",1);
    cellDescription.addParagraph(getValue(param, "description", "chinese"),"");
    if (userParam.printcambodia) {
        cellDescription.addParagraph(getValue(param, "description", "cambodia"),"");
    }

    var cellGenLedAc = tableRow.addCell("","alignCenter heightTrans",1);
    cellGenLedAc.addParagraph(getValue(param, "genLedAc", "chinese"),"");
    if (userParam.printcambodia) {
        cellGenLedAc.addParagraph(getValue(param, "genLedAc", "cambodia"),"");
    }

    var cellDebitAmount = tableRow.addCell("","alignCenter heightTrans",1);
    cellDebitAmount.addParagraph(getValue(param, "debitAmount", "chinese"),"");
    if (userParam.printcambodia) {
        cellDebitAmount.addParagraph(getValue(param, "debitAmount", "cambodia"),"");
    }

    var cellCreditAmount = tableRow.addCell("","alignCenter heightTrans",1);
    cellCreditAmount.addParagraph(getValue(param, "creditAmount", "chinese"),"");
    if (userParam.printcambodia) {
        cellCreditAmount.addParagraph(getValue(param, "creditAmount", "cambodia"),"");
    }

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
                    tableRow.addCell(tRow.value('Description'), "text-black padding-left heightTrans", 1);
                } else {
                    tableRow.addCell("", "heightTrans", 1);
                }
                tableRow.addCell(tRow.value('JAccountDescription'), "text-black padding-left heightTrans", 1);
                
                // Debit
                if (Banana.SDecimal.sign(tRow.value('JAmount')) > 0 ) {
                    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amount), "text-black padding-left padding-right alignRight heightTrans", 1);
                    tableRow.addCell("", "heightTrans", 1);
                }
                // Credit
                else {
                    tableRow.addCell("", "heightTrans", 1);
                    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(amount), "text-black padding-left padding-right alignRight heightTrans", 1);
                }

                line++;
            }
        }
    }

    //Add empty lines if current lines are not enough
    while (line < generalParam.numberOfRows) {
        tableRow = table.addRow();
        tableRow.addCell("", "heightTrans", 1);
        tableRow.addCell("", "heightTrans", 1);
        tableRow.addCell("", "heightTrans", 1);
        tableRow.addCell("", "heightTrans", 1);
        line++;
    }
}

/* Function that prints the total line of the voucher */
function printTotal(table, totDebit, totCredit, report, userParam) {

    //Attachments
    tableRow = table.addRow();
    if (userParam.printcambodia) {
        tableRow.addCell(getValue(param, "attachments", "chinese") + "   " + getValue(param, "attachments", "cambodia"), "alignCenter padding-left heightTotal", 1);
    } else {
        tableRow.addCell(getValue(param, "attachments", "chinese"), "alignCenter padding-left heightTotal", 1);
    }

    //Total
    if (userParam.printcambodia) {
        tableRow.addCell(getValue(param, "total", "chinese") + "   " + getValue(param, "total", "cambodia"), "alignCenter padding-left heightTotal", 1);
    } else {
        tableRow.addCell(getValue(param, "total", "chinese"), "alignCenter padding-left heightTotal", 1);
    }

    if (totDebit === "" || totDebit === "undefined" || !totDebit) {
        totDebit = "0.00";
    }
    if (totCredit === "" || totCredit === "undefined" || !totCredit) {
        totCredit = "0.00";
    }

    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(totDebit), "text-black padding-left padding-right alignRight heightTotal",1);
    tableRow.addCell(Banana.Converter.toLocaleNumberFormat(totCredit), "text-black padding-left padding-right alignRight heightTotal",1);
}

/* Function that prints the signature part of the voucher */
function printSignatures(tableSignatures, report, userParam) {
    var paragraph;
    tableRow = tableSignatures.addRow();
    var cellApproved = tableRow.addCell("", "", 1);
    paragraph = cellApproved.addParagraph();
    paragraph.addText(getValue(param, "approved", "chinese") + ": ", "");
    paragraph.addText(userParam.approved, "text-black");
    if (userParam.printcambodia) {
        paragraph = cellApproved.addParagraph();
        paragraph.addText(getValue(param, "approved", "cambodia"), "");
    }
    
    var cellChecked = tableRow.addCell("", "", 1);
    paragraph = cellChecked.addParagraph();
    paragraph.addText(getValue(param, "checked", "chinese") + ": ", "");
    paragraph.addText(userParam.checked, "text-black");
    if (userParam.printcambodia) {
        paragraph = cellChecked.addParagraph();
        paragraph.addText(getValue(param, "checked", "cambodia"), "");
    }

    var cellEntered = tableRow.addCell("", "", 1);
    paragraph = cellEntered.addParagraph();
    paragraph.addText(getValue(param, "entered", "chinese") + ": ", "");
    paragraph.addText(userParam.entered, "text-black");
    if (userParam.printcambodia) {
        paragraph = cellEntered.addParagraph();
        paragraph.addText(getValue(param, "entered", "cambodia"), "");
    }

    var cellCashier = tableRow.addCell("", "", 1);
    paragraph = cellCashier.addParagraph();
    paragraph.addText(getValue(param, "cashier", "chinese") + ": ", "");
    paragraph.addText(userParam.cashier, "text-black");
    if (userParam.printcambodia) {
        paragraph = cellCashier.addParagraph();
        paragraph.addText(getValue(param, "cashier", "cambodia"), "");
    }

    var cellPrepared = tableRow.addCell("", "", 1);
    paragraph = cellPrepared.addParagraph();
    paragraph.addText(getValue(param, "prepared", "chinese") + ": ", "");
    paragraph.addText(userParam.prepared, "text-black");
    if (userParam.printcambodia) {
        paragraph = cellPrepared.addParagraph();
        paragraph.addText(getValue(param, "prepared", "cambodia"), "");
    }

    var cellReceiver = tableRow.addCell("", "", 1);
    paragraph = cellReceiver.addParagraph();
    paragraph.addText(getValue(param, "receiver", "chinese") + ": ", "");
    paragraph.addText(userParam.receiver, "text-black");
    if (userParam.printcambodia) {
        paragraph = cellReceiver.addParagraph();
        paragraph.addText(getValue(param, "receiver", "cambodia"), "");
    }
}

/* Function that creates a list with all the values taken from the columna "Doc" of the table "Transactions" */
function getDocList(banDoc) {
    var str = [];
    for (var i = 0; i < banDoc.table('Transactions').rowCount; i++) {
        var tRow = banDoc.table('Transactions').row(i);
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
    pageStyle.setAttribute("margin", "3mm 5mm 3mm 5mm");
    pageStyle.setAttribute("size", generalParam.pageSize);

    

    /*
        General styles
    */
    stylesheet.addStyle("body", "font-family : Times New Roman; color:" + generalParam.textColor);
    stylesheet.addStyle(".text-black", "color:black");
    stylesheet.addStyle(".border-top-double", "border-top: 0.8px double black");

    stylesheet.addStyle(".padding-left", "padding-left:5px");
    stylesheet.addStyle(".padding-right", "padding-right:5px");
    stylesheet.addStyle(".underLine", "border-top:thin double black");

    stylesheet.addStyle(".heading1", "font-size:"+generalParam.fontSize+";font-weight:bold");
    stylesheet.addStyle(".heading2", "font-size:"+generalParam.fontSize+";");
    stylesheet.addStyle(".heading3", "font-size:"+generalParam.fontSize+";");
    stylesheet.addStyle(".bold", "font-weight:bold");
    stylesheet.addStyle(".alignRight", "text-align:right");
    stylesheet.addStyle(".alignCenter", "text-align:center");

 
    /* 
        Info table style
    */
    style = stylesheet.addStyle("table_info");
    style.setAttribute("width", "100%");
    style.setAttribute("font-size", generalParam.fontSize);
    stylesheet.addStyle("table.table_info td", "padding-bottom:2px;");

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
    style.setAttribute("font-size", generalParam.fontSizeTable);
    stylesheet.addStyle("table.table td", "border:thin solid " + generalParam.textColor+";");
    // stylesheet.addStyle("table.table tr", "height: " + generalParam.rowHeight+";");
    stylesheet.addStyle(".heightTrans", "height: " + generalParam.rowHeight+";");
    stylesheet.addStyle(".heightTotal", "height: " + generalParam.rowHeightTotal+";");

    //Columns for the transactions table
    stylesheet.addStyle(".c1", "width:"+generalParam.column1);
    stylesheet.addStyle(".c2", "width:"+generalParam.column2);
    stylesheet.addStyle(".c3", "width:"+generalParam.column3);
    stylesheet.addStyle(".c4", "width:"+generalParam.column4);


    /*
        Signatures table style
    */
    style = stylesheet.addStyle("table_signatures");
    style.setAttribute("width", "100%");
    style.setAttribute("font-size", generalParam.fontSize);
    stylesheet.addStyle("table.table_signatures td", "");

    //Column for the signatures table
    stylesheet.addStyle(".colSig1", "width:16.6%");
    stylesheet.addStyle(".colSig2", "width:16.6%");
    stylesheet.addStyle(".colSig3", "width:16.6%");
    stylesheet.addStyle(".colSig4", "width:16.6%");
    stylesheet.addStyle(".colSig5", "width:16.6%");
    stylesheet.addStyle(".colSig6", "width:16.6%");

    return stylesheet;
}

/* Function that converts parameters of the dialog */
function convertParam(userParam) {

    var convertedParam = {};
    convertedParam.version = '1.0';
    convertedParam.data = []; /* array dei parametri dello script */


    //Text to explain
    var currentParam = {};
    currentParam.name = 'explaintext1';
    currentParam.title = '单击 "数值" 下的空白区域，输入需打印的凭证号码';
    currentParam.type = 'string';
    currentParam.value = '';
    currentParam.editable = false;
    currentParam.readValue = function() {
        userParam.explaintext1 = this.value;
    }
    convertedParam.data.push(currentParam);

    //Text to explain
    var currentParam = {};
    currentParam.name = 'explaintext2';
    currentParam.title = '(输入“*”号可一次性打印所有凭证) 及相关人员名称 。';
    currentParam.type = 'string';
    currentParam.value = '';
    currentParam.editable = false;
    currentParam.readValue = function() {
        userParam.explaintext2 = this.value;
    }
    convertedParam.data.push(currentParam);


    // parameters
    var currentParam = {};
    currentParam.name = 'parameters';
    currentParam.title = '信息填写';
    currentParam.type = 'string';
    currentParam.value = '';
    currentParam.editable = false;
    currentParam.readValue = function() {
        userParam.parameters = this.value;
    }
    convertedParam.data.push(currentParam);

    //Voucher number
    var currentParam = {};
    currentParam.name = 'voucher';
    currentParam.parentObject = 'parameters';
    currentParam.title = "记账凭证号码:";
    currentParam.type = 'string';
    
    //Get the name of the current selected table
    //If the current table is the Transactions table, we take the value of the column "Doc" of the selected row
    var currentTable = Banana.document.cursor.tableName;
    if (currentTable === "Transactions") {
        var transactions = Banana.document.table('Transactions');
        var tDoc = transactions.row(Banana.document.cursor.rowNr).value('Doc');
        currentParam.value = tDoc;
    } else {
        currentParam.value = '';
    }
    currentParam.readValue = function() {
        userParam.voucher = this.value;
    }
    convertedParam.data.push(currentParam);

    //Approved
    var currentParam = {};
    currentParam.name = 'approved';
    currentParam.parentObject = 'parameters';
    currentParam.title = '核准:';
    currentParam.type = 'string';
    currentParam.value = userParam.approved ? userParam.approved : '';
    currentParam.readValue = function() {
        userParam.approved = this.value;
    }
    convertedParam.data.push(currentParam);

    //Checked
    var currentParam = {};
    currentParam.name = 'checked';
    currentParam.parentObject = 'parameters';
    currentParam.title = '复核:';
    currentParam.type = 'string';
    currentParam.value = userParam.checked ? userParam.checked : '';
    currentParam.readValue = function() {
        userParam.checked = this.value;
    }
    convertedParam.data.push(currentParam);

    //Entered
    var currentParam = {};
    currentParam.name = 'entered';
    currentParam.parentObject = 'parameters';
    currentParam.title = '记账:';
    currentParam.type = 'string';
    currentParam.value = userParam.entered ? userParam.entered : '';
    currentParam.readValue = function() {
        userParam.entered = this.value;
    }
    convertedParam.data.push(currentParam);

    //Cashier
    var currentParam = {};
    currentParam.name = 'cashier';
    currentParam.parentObject = 'parameters';
    currentParam.title = '出纳:';
    currentParam.type = 'string';
    currentParam.value = userParam.cashier ? userParam.cashier : '';
    currentParam.readValue = function() {
        userParam.cashier = this.value;
    }
    convertedParam.data.push(currentParam);

    //Prepared
    var currentParam = {};
    currentParam.name = 'prepared';
    currentParam.parentObject = 'parameters';
    currentParam.title = '制单:';
    currentParam.type = 'string';
    currentParam.value = userParam.prepared ? userParam.prepared : '';
    currentParam.readValue = function() {
        userParam.prepared = this.value;
    }
    convertedParam.data.push(currentParam);

    //Receiver
    var currentParam = {};
    currentParam.name = 'receiver';
    currentParam.parentObject = 'parameters';
    currentParam.title = '签收:';
    currentParam.type = 'string';
    currentParam.value = userParam.receiver ? userParam.receiver : '';
    currentParam.readValue = function() {
        userParam.receiver = this.value;
    }
    convertedParam.data.push(currentParam);

    // print output
    var currentParam = {};
    currentParam.name = 'printoutput';
    currentParam.title = '凭证打印设置';
    currentParam.type = 'string';
    currentParam.value = '';
    currentParam.editable = false;
    currentParam.readValue = function() {
        userParam.printoutput = this.value;
    }
    convertedParam.data.push(currentParam);

    // print cambodia translations
    var currentParam = {};
    currentParam.name = 'printcambodia';
    currentParam.parentObject = 'printoutput'
    currentParam.title = '用中文和高棉语打印凭单（如果您不选中此选项，则默认情况下将打印为中文凭单）';
    currentParam.type = 'bool';
    currentParam.value = userParam.printcambodia ? true : false;
    currentParam.readValue = function() {
     userParam.printcambodia = this.value;
    }
    convertedParam.data.push(currentParam);


    // Page size
    var currentParam = {};
    currentParam.name = 'pagesize';
    currentParam.title = '页面设置';
    currentParam.type = 'string';
    currentParam.value = '';
    currentParam.editable = false;
    currentParam.readValue = function() {
        userParam.pagesize = this.value;
    }
    convertedParam.data.push(currentParam);

    // 240mm x 140mm
    var currentParam = {};
    currentParam.name = 'pagesize0';
    currentParam.parentObject = 'pagesize'
    currentParam.title = '宽度:240 毫米; 高度:140 毫米';
    currentParam.type = 'bool';
    currentParam.value = userParam.pagesize0 ? true : false;
    currentParam.readValue = function() {
     userParam.pagesize0 = this.value;
    }
    convertedParam.data.push(currentParam);

    // 243mm x 142mm
    var currentParam = {};
    currentParam.name = 'pagesize1';
    currentParam.parentObject = 'pagesize'
    currentParam.title = '宽度:243 毫米; 高度:142 毫米';
    currentParam.type = 'bool';
    currentParam.value = userParam.pagesize1 ? true : false;
    currentParam.readValue = function() {
     userParam.pagesize1 = this.value;
    }
    convertedParam.data.push(currentParam);

    // 210mm x 127mm
    var currentParam = {};
    currentParam.name = 'pagesize2';
    currentParam.parentObject = 'pagesize'
    currentParam.title = '宽度:210 毫米; 高度:127 毫米';
    currentParam.type = 'bool';
    currentParam.value = userParam.pagesize2 ? true : false;
    currentParam.readValue = function() {
     userParam.pagesize2 = this.value;
    }
    convertedParam.data.push(currentParam);

    // A5 landscape
    var currentParam = {};
    currentParam.name = 'pagesize3';
    currentParam.parentObject = 'pagesize'
    currentParam.title = 'A5 横向';
    currentParam.type = 'bool';
    currentParam.value = userParam.pagesize3 ? true : false;
    currentParam.readValue = function() {
     userParam.pagesize3 = this.value;
    }
    convertedParam.data.push(currentParam);

    // A4
    var currentParam = {};
    currentParam.name = 'pagesize4';
    currentParam.parentObject = 'pagesize'
    currentParam.title = 'A4 纵向';
    currentParam.type = 'bool';
    currentParam.value = userParam.pagesize4 ? true : false;
    currentParam.readValue = function() {
     userParam.pagesize4 = this.value;
     if (userParam.pagesize0 && (userParam.pagesize1 || userParam.pagesize2 || userParam.pagesize3 || userParam.pagesize4)) {
        userParam.pagesize0 = false;
        userParam.pagesize1 = false;
        userParam.pagesize2 = false;
        userParam.pagesize3 = false;
        userParam.pagesize4 = true;
     }     
     if (userParam.pagesize1 && (userParam.pagesize0 || userParam.pagesize2 || userParam.pagesize3 || userParam.pagesize4)) {
        userParam.pagesize0 = false;
        userParam.pagesize1 = false;
        userParam.pagesize2 = false;
        userParam.pagesize3 = false;
        userParam.pagesize4 = true;
     }
     if (userParam.pagesize2 && (userParam.pagesize0 || userParam.pagesize1 || userParam.pagesize3 || userParam.pagesize4)) {
        userParam.pagesize0 = false;
        userParam.pagesize1 = false;
        userParam.pagesize2 = false;
        userParam.pagesize3 = false;
        userParam.pagesize4 = true;
     }
     if (userParam.pagesize3 && (userParam.pagesize0 || userParam.pagesize1 || userParam.pagesize2 || userParam.pagesize4)) {
        userParam.pagesize0 = false;
        userParam.pagesize1 = false;
        userParam.pagesize2 = false;
        userParam.pagesize3 = false;
        userParam.pagesize4 = true;
     }
     if (userParam.pagesize4) {
        userParam.pagesize0 = false;
        userParam.pagesize1 = false;
        userParam.pagesize2 = false;
        userParam.pagesize3 = false;
     }
     if (!userParam.pagesize0 && !userParam.pagesize1 && !userParam.pagesize2 && !userParam.pagesize3 && !userParam.pagesize4 && !userParam.custompagesize) {
        userParam.pagesize4 = true;
     }
    }
    convertedParam.data.push(currentParam);

    // Custom page size
    var currentParam = {};
    currentParam.name = 'custompagesize';
    currentParam.parentObject = 'pagesize';
    currentParam.title = '自定义大小';
    currentParam.type = 'bool';
    currentParam.value = userParam.custompagesize ? true : false;
    currentParam.readValue = function() {
        userParam.custompagesize = this.value;
        if (userParam.custompagesize) {
            userParam.pagesize0 = false;
            userParam.pagesize1 = false;
            userParam.pagesize2 = false;
            userParam.pagesize3 = false;
            userParam.pagesize4 = false;
        }
    }
    convertedParam.data.push(currentParam);

    // custom width size
    var currentParam = {};
    currentParam.name = 'customwidth';
    currentParam.parentObject = 'custompagesize';
    currentParam.title = '自定义宽度 (毫米)';
    currentParam.type = 'string';
    currentParam.value = userParam.customwidth ? userParam.customwidth : '';
    currentParam.readValue = function() {
        userParam.customwidth = this.value;
    }
    convertedParam.data.push(currentParam);

    // custom height size
    var currentParam = {};
    currentParam.name = 'customheight';
    currentParam.parentObject = 'custompagesize';
    currentParam.title = '自定义高度 (毫米)';
    currentParam.type = 'string';
    currentParam.value = userParam.customheight ? userParam.customheight : '';
    currentParam.readValue = function() {
        userParam.customheight = this.value;
    }
    convertedParam.data.push(currentParam);

    // custom landscape choice
    var currentParam = {};
    currentParam.name = 'landscape';
    currentParam.parentObject = 'custompagesize'
    currentParam.title = '横向';
    currentParam.type = 'bool';
    currentParam.value = userParam.landscape ? true : false;
    currentParam.readValue = function() {
     userParam.landscape = this.value;
    }
    convertedParam.data.push(currentParam);

    // Style
    var currentParam = {};
    currentParam.name = 'style';
    currentParam.title = '字体颜色';
    currentParam.type = 'string';
    currentParam.value = '';
    currentParam.editable = false;
    currentParam.readValue = function() {
        userParam.style = this.value;
    }
    convertedParam.data.push(currentParam);

    // Color
    var currentParam = {};
    currentParam.name = 'color';
    currentParam.parentObject = 'style';
    currentParam.title = '颜色';
    currentParam.type = 'string';
    currentParam.value = userParam.color ? userParam.color : 'Black';
    currentParam.readValue = function() {
        userParam.color = this.value;
    }
    convertedParam.data.push(currentParam);


    return convertedParam;
}

/* Function that initializes the user parameters */
function initUserParam() {
    var userParam = {};
    userParam.version = '1.0';
    userParam.explaintext1 = '';
    userParam.explaintext2 = '';
    userParam.parameters = '';
    userParam.voucher = '';
    userParam.approved = '';
    userParam.checked = '';
    userParam.entered = '';
    userParam.cashier = '';
    userParam.prepared = '';
    userParam.receiver = '';
    userParam.printoutput = '';
    userParam.printcambodia = true;
    userParam.pagesize0 = false;
    userParam.pagesize1 = false;
    userParam.pagesize2 = false;
    userParam.pagesize3 = false;
    userParam.pagesize4 = true;
    userParam.custompagesize = false;
    userParam.customwidth = '';
    userParam.customheight = '';
    userParam.landscape = false;
    userParam.pagesize5 = '';
    userParam.color = 'Black';

    return userParam;
}

/* Function that shows the dialog window and let user to modify the parameters */
function parametersDialog(userParam) {

    if (typeof(Banana.Ui.openPropertyEditor) !== 'undefined') {
        var dialogTitle = '设置';
        var convertedParam = convertParam(userParam);
        var pageAnchor = 'dlgSettings';
        if (!Banana.Ui.openPropertyEditor(dialogTitle, convertedParam, pageAnchor)) {
            return null;
        }
        
        for (var i = 0; i < convertedParam.data.length; i++) {
            // Read values to userParam (through the readValue function)
            convertedParam.data[i].readValue();
        }
    }
    
    return userParam;
}

/* Function that shows a dialog window for the period and let user to modify the parameters */
function settingsDialog() {

    var scriptform = initUserParam();
    
    // Retrieve saved param
    var savedParam = Banana.document.getScriptSettings();
    if (savedParam && savedParam.length > 0) {
        scriptform = JSON.parse(savedParam);
    }

    scriptform = parametersDialog(scriptform); // From propertiess
    if (scriptform) {
        var paramToString = JSON.stringify(scriptform);
        Banana.document.setScriptSettings(paramToString);
    }

    return scriptform;
}
