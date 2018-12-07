// Copyright [2018] [Banana.ch SA - Lugano Switzerland]
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
// @id = ch.banana.china.reportprofitloss
// @api = 1.0
// @pubdate = 2018-12-07
// @publisher = Banana.ch SA
// @description.zh = 利润表
// @description.en = Profit and Loss
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @outputformat = none
// @inputdataform = none
// @timeout = -1

function loadParam(banDoc, startDate, endDate) {
	var date = new Date();
	var param = {
		"reportName":"China - Profit/Loss Report",
		"startDate":startDate,
		"endDate":endDate,
		"currentDate": date,
		"taxpayerNumber":banDoc.info("AccountingDataBase","FiscalNumber"),
		"company":banDoc.info("AccountingDataBase","Company"),
		"address":banDoc.info("AccountingDataBase","Address1"),
		"nation":banDoc.info("AccountingDataBase","Country"),
		"telephone":banDoc.info("AccountingDataBase","Phone"),
		"zip":banDoc.info("AccountingDataBase","Zip"),
		"city":banDoc.info("AccountingDataBase","City")
	};
	return param;
}

function exec(string) {

	if (!Banana.document) {
		return "@Cancel";
	}

	var dateform = getPeriodSettings();
	if (dateform) {
		var report = createProfitReport(Banana.document, dateform.selectionStartDate, dateform.selectionEndDate);
		var stylesheet = createStylesheet();
		Banana.Report.preview(report, stylesheet);	
	} else {
		return "@Cancel";
	}
}

function createProfitReport(banDoc, startDate, endDate) {

	var param = loadParam(banDoc, startDate, endDate);
	var report = Banana.Report.newReport(param.reportName);

	report.addParagraph("利润表(适用执行小企业会计准则的企业)", "bold"); //Balance sheet (application of accounting standards for small business enterprise)
	report.addParagraph(" ", "");

	var table = report.addTable("table");
	var col1 = table.addColumn("col1");
	var col2 = table.addColumn("col2");
	var col3 = table.addColumn("col3");
	var col4 = table.addColumn("col4");
	var col5 = table.addColumn("col5");
	var col6 = table.addColumn("col6");
	var col7 = table.addColumn("col7");
	var col8 = table.addColumn("col8");

	var amount,monthamount = "";
	var totAmount13,totAmountMonth13 = "";
	var totAmount18,totAmountMonth18 = "";
	var totAmount20,totAmountMonth20 = "";


	tableRow = table.addRow();
	tableRow.addCell("纳税人基本信息", "center border-bottom", 4); //Taxpayer information

	tableRow = table.addRow();
	tableRow.addCell("", "", 4);
	
	tableRow = table.addRow();
	tableRow.addCell("纳税人识别号: " + param["taxpayerNumber"], "", 2); //Taxpayer identification number:
	tableRow.addCell("纳税人名称: " + param["company"], "", 2); //Payer's name

	tableRow = table.addRow();
	var d = new Date();
	tableRow.addCell("报送日期: " + Banana.Converter.toLocaleDateFormat(d), "", 2); //Submit date
	tableRow.addCell("所属时期起: " + Banana.Converter.toLocaleDateFormat(startDate), "", 1); //Period
	tableRow.addCell("所属时期止: " + Banana.Converter.toLocaleDateFormat(endDate), "", 1); //Period

	tableRow = table.addRow();
	tableRow.addCell("", "", 4);

	tableRow = table.addRow();	
	tableRow.addCell("项   目", "valueTitle bold bordersTitle border-left border-top border-bottom center", 1); //Project
	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom center", 1); //Line no
	tableRow.addCell("本期金额", "valueTitle bold bordersTitle border-top border-bottom center", 1); //this month amount
	tableRow.addCell("本年累计金额", "valueTitle bold bordersTitle border-top border-bottom center", 1); //YTD amount

	// 1
	tableRow = table.addRow();
	tableRow.addCell("一、营业收入", "borders", 1);
	tableRow.addCell("1", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6001|6041|6051", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total);
	monthamount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6001|6041|6051", param["startDate"], param["endDate"]).total);
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.add(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.add(totAmountMonth13,monthamount);
	
	// 2
	tableRow = table.addRow();
	tableRow.addCell("减：营业成本", "borders", 1);
	tableRow.addCell("2", "borders", 1);
	amount = banDoc.currentBalance("Gr=6401|6402", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6401|6402", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.subtract(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.subtract(totAmountMonth13,monthamount);
	
	// 3
	tableRow = table.addRow();
	tableRow.addCell("税金及附加", "borders", 1);
	tableRow.addCell("3", "borders", 1);
	amount = banDoc.currentBalance("Gr=6403", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6403", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.subtract(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.subtract(totAmountMonth13,monthamount);	
	
	// 4
	tableRow = table.addRow();
	tableRow.addCell("销售费用", "borders", 1);
	tableRow.addCell("4", "borders", 1);
	amount = banDoc.currentBalance("Gr=6601", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6601", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.subtract(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.subtract(totAmountMonth13,monthamount);
	
	// 5
	tableRow = table.addRow();
	tableRow.addCell("管理费用", "borders", 1);
	tableRow.addCell("5", "borders", 1);
	amount = banDoc.currentBalance("Gr=6602", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6602", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.subtract(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.subtract(totAmountMonth13,monthamount);	
	
	// 6
	tableRow = table.addRow();
	tableRow.addCell("财务费用", "borders", 1);
	tableRow.addCell("6", "borders", 1);
	amount = banDoc.currentBalance("Gr=6603", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6603", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.subtract(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.subtract(totAmountMonth13,monthamount);	

	// 7
	tableRow = table.addRow();
	tableRow.addCell("资产减值损失", "borders", 1);
	tableRow.addCell("7", "borders", 1);
	amount = banDoc.currentBalance("Gr=6701", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6701", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.subtract(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.subtract(totAmountMonth13,monthamount);

	// 8
	tableRow = table.addRow();
	tableRow.addCell("加：公允价值变动收益（损失以“-”号填列）", "borders", 1);
	tableRow.addCell("8", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6101", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total);
	monthamount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6101", param["startDate"], param["endDate"]).total);
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.add(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.add(totAmountMonth13,monthamount);

	// 9
	tableRow = table.addRow();
	tableRow.addCell("投资收益（损失以“-”号填列）", "borders", 1);
	tableRow.addCell("9", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6111", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total);
	monthamount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6111", param["startDate"], param["endDate"]).total);
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount13 = Banana.SDecimal.add(totAmount13,amount);
	totAmountMonth13 = Banana.SDecimal.add(totAmountMonth13,monthamount);

	// 10
	tableRow = table.addRow();
	tableRow.addCell("其中：对联营企业和合营企业的投资收益", "borders", 1);
	tableRow.addCell("10", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);

	// 11
	tableRow = table.addRow();
	tableRow.addCell("资产处置收益（损失以“-”号填列）", "borders", 1);
	tableRow.addCell("11", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);

	// 12
	tableRow = table.addRow();
	tableRow.addCell("其他收益", "borders", 1);
	tableRow.addCell("12", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);

	// 13 => total: 1-2-3-4-5-6-7+8+9
	tableRow = table.addRow();
	tableRow.addCell("二、营业利润（亏损以“-”号填列）", "borders styleTotal", 1);
	tableRow.addCell("13", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmountMonth13), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totAmount13), "borders right padding-right styleTotal", 1);
	totAmount18 = Banana.SDecimal.add(totAmount18,totAmount13);
	totAmountMonth18 = Banana.SDecimal.add(totAmountMonth18,totAmountMonth13);

	// 14
	tableRow = table.addRow();
	tableRow.addCell("加：营业外收入", "borders", 1);
	tableRow.addCell("14", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6301", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total);
	monthamount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=6301", param["startDate"], param["endDate"]).total);
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount18 = Banana.SDecimal.add(totAmount18,amount);
	totAmountMonth18 = Banana.SDecimal.add(totAmountMonth18,monthamount);

	// 15
	tableRow = table.addRow();
	tableRow.addCell("其中：非流动资产处置利得", "borders", 1);
	tableRow.addCell("15", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);

	// 16
	tableRow = table.addRow();
	tableRow.addCell("减：营业外支出", "borders", 1);
	tableRow.addCell("16", "borders", 1);
	amount = banDoc.currentBalance("Gr=6711", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6711", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount18 = Banana.SDecimal.subtract(totAmount18,amount);
	totAmountMonth18 = Banana.SDecimal.subtract(totAmountMonth18,monthamount);

	// 17
	tableRow = table.addRow();
	tableRow.addCell("其中：非流动资产处置损失", "borders", 1);
	tableRow.addCell("17", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);

	// 18 => total: 13+14-16
	tableRow = table.addRow();
	tableRow.addCell("三、利润总额（亏损总额以“-”号填列）", "borders styleTotal", 1);
	tableRow.addCell("18", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmountMonth18), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totAmount18), "borders right padding-right styleTotal", 1);
	totAmount20 = Banana.SDecimal.add(totAmount20,totAmount18);
	totAmountMonth20 = Banana.SDecimal.add(totAmountMonth20,totAmountMonth18);

	// 19
	tableRow = table.addRow();
	tableRow.addCell("减：所得税费用", "borders", 1);
	tableRow.addCell("19", "borders", 1);
	amount = banDoc.currentBalance("Gr=6801", banDoc.info("AccountingDataBase","OpeningDate"), param["endDate"]).total;
	monthamount = banDoc.currentBalance("Gr=6801", param["startDate"], param["endDate"]).total;
	tableRow.addCell(formatValues(monthamount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	totAmount20 = Banana.SDecimal.subtract(totAmount20,amount);
	totAmountMonth20 = Banana.SDecimal.subtract(totAmountMonth20,monthamount);

	// 20 => total: 18-19
	tableRow = table.addRow();
	tableRow.addCell("四、净利润（净亏损以“-”号填列）", "borders styleTotal", 1);
	tableRow.addCell("20", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmountMonth20), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totAmount20), "borders right padding-right styleTotal", 1);
	
	return report;
}

function formatValues(value) {
	if (!value || value === "0" || value == null) {
		value = "0";
	}
	return Banana.Converter.toLocaleNumberFormat(value);
}

function getPeriodSettings() {
	
	//The formeters of the period that we need
	var scriptform = {
	   "selectionStartDate": "",
	   "selectionEndDate": "",
	   "selectionChecked": "false"
	};

	//Read script settings
	var data = Banana.document.scriptReadSettings();
	
	//Check if there are previously saved settings and read them
	if (data.length > 0) {
		try {
			var readSettings = JSON.parse(data);
			
			//We check if "readSettings" is not null, then we fill the formeters with the values just read
			if (readSettings) {
				scriptform = readSettings;
			}
		} catch (e){}
	}
	
	//We take the accounting "starting date" and "ending date" from the document. These will be used as default dates
	var docStartDate = Banana.document.startPeriod();
	var docEndDate = Banana.document.endPeriod();	
	
	//A dialog window is opened asking the user to insert the desired period. By default is the accounting period
	var selectedDates = Banana.Ui.getPeriod("Period", docStartDate, docEndDate, 
		scriptform.selectionStartDate, scriptform.selectionEndDate, scriptform.selectionChecked);
		
	//We take the values entered by the user and save them as "new default" values.
	//This because the next time the script will be executed, the dialog window will contains the new values.
	if (selectedDates) {
		scriptform["selectionStartDate"] = selectedDates.startDate;
		scriptform["selectionEndDate"] = selectedDates.endDate;
		scriptform["selectionChecked"] = selectedDates.hasSelection;

		//Save script settings
		var formToString = JSON.stringify(scriptform);
		var value = Banana.document.scriptSaveSettings(formToString);		
    } else {
		//User clicked cancel
		return;
	}
	return scriptform;
}

function createStylesheet() {
	var stylesheet = Banana.Report.newStyleSheet();

    var pageStyle = stylesheet.addStyle("@page");
    pageStyle.setAttribute("margin", "10mm 10mm 10mm 20mm");

    stylesheet.addStyle("body", "font-size: 9pt; font-family: Helvetica");
    stylesheet.addStyle(".bold", "font-weight:bold");
    stylesheet.addStyle(".center", "text-align:center");
    stylesheet.addStyle(".right", "text-align:right");

    stylesheet.addStyle(".border-left", "border-left:thin solid black");
    stylesheet.addStyle(".border-right", "border-right:thin solid black");
    stylesheet.addStyle(".border-top", "border-top:thin solid black");
    stylesheet.addStyle(".border-bottom", "border-bottom:thin solid black");
    stylesheet.addStyle(".padding-right", "padding-right:2px");

    stylesheet.addStyle(".col1", "width:%");
	stylesheet.addStyle(".col2", "width:%");
	stylesheet.addStyle(".col3", "width:%");
	stylesheet.addStyle(".col4", "width:%");
	stylesheet.addStyle(".col5", "width:%");
	stylesheet.addStyle(".col6", "width:%");
	stylesheet.addStyle(".col7", "width:%");
	stylesheet.addStyle(".col8", "width:%");

	style = stylesheet.addStyle(".valueTitle");
	style.setAttribute("background-color", "#000000");
	style.setAttribute("color", "#fff");

	style = stylesheet.addStyle(".bordersTitle");
	style.setAttribute("border-right", "thin solid #fff");
	style.setAttribute("padding-bottom", "2px");
	style.setAttribute("padding-top", "5px"); 

	style = stylesheet.addStyle(".borders");
	style.setAttribute("border-left", "thin solid black");
	style.setAttribute("border-right", "thin solid black");
	style.setAttribute("border-bottom", "thin solid black");
	style.setAttribute("border-top", "thin solid black");

	style = stylesheet.addStyle(".table");
	style.setAttribute("width", "100%");
	//stylesheet.addStyle("table.table td", "border: thin solid black");

    style = stylesheet.addStyle(".styleTotal");
    style.setAttribute("font-weight", "bold");
    style.setAttribute("background-color", "#b7c3e0");

	return stylesheet;
}
