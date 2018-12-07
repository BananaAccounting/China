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
// @id = ch.banana.china.reportbalancesheet
// @api = 1.0
// @pubdate = 2018-12-07
// @publisher = Banana.ch SA
// @description.zh = 资产负债表
// @description.en = Balance Sheet
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @outputformat = none
// @inputdataform = none
// @timeout = -1

function loadParam(banDoc, startDate, endDate) {
	var param = {
		"reportName":"China - Balance Report",
		"startDate":startDate,
		"endDate":endDate,
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
		var report = createBalanceReport(Banana.document, dateform.selectionStartDate, dateform.selectionEndDate);
		var stylesheet = createStylesheet();
		Banana.Report.preview(report, stylesheet);	
	} else {
		return "@Cancel";
	}
}

function createBalanceReport(banDoc, startDate, endDate) {

	var param = loadParam(banDoc, startDate, endDate);
	var report = Banana.Report.newReport(param.reportName);

	report.addParagraph("资产负债表(适用执行小企业会计准则的企业)", "bold"); //Balance sheet (application of accounting standards for small business enterprise)

	var table = report.addTable("table");
	var col1 = table.addColumn("col1");
	var col2 = table.addColumn("col2");
	var col3 = table.addColumn("col3");
	var col4 = table.addColumn("col4");
	var col5 = table.addColumn("col5");
	var col6 = table.addColumn("col6");
	var col7 = table.addColumn("col7");
	var col8 = table.addColumn("col8");
	
	var amount,opening = "";
	var totAmount12,totOpening12 = "";
	var totAmount30,totOpening30 = "";
	var totAmount44,totOpening44 = "";
	var totAmount52,totOpening52 = "";
	var totAmount53,totOpening53 = "";
	var totAmount61,totOpening61 = "";
	var totAmount31,totOpening31 = "";
	var totAmount62,totOpening62 = "";

	tableRow = table.addRow();
	tableRow.addCell("纳税人基本信息", "center border-bottom", 8); //Taxpayer information
	
	tableRow = table.addRow();
	tableRow.addCell("", "", 8);
	
	tableRow = table.addRow();
	tableRow.addCell("纳税人识别号: " + param["taxpayerNumber"], "", 4); //Taxpayer identification number:
	tableRow.addCell("纳税人名称: " + param["company"], "", 4); //Payer's name

	tableRow = table.addRow();
	var d = new Date();
	tableRow.addCell("报送日期: " + Banana.Converter.toLocaleDateFormat(d), "", 4); //Submit date
	tableRow.addCell("所属时期起: " + Banana.Converter.toLocaleDateFormat(startDate), "", 1); //Period
	tableRow.addCell("所属时期止: " + Banana.Converter.toLocaleDateFormat(endDate), "", 3); //Period

	tableRow = table.addRow();
	tableRow.addCell("", "", 8);

	tableRow = table.addRow();	
	tableRow.addCell("资 产", "valueTitle bold bordersTitle border-left border-top border-bottom", 1); 	//Assets
	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom", 1); 				//Line no
	tableRow.addCell("期末余额", "valueTitle bold bordersTitle border-top border-bottom", 1); 			//Ending balance
	tableRow.addCell("年初余额", "valueTitle bold bordersTitle border-top border-bottom", 1); 			//At beginning of year
	tableRow.addCell("负债及所有者权益", "valueTitle bold bordersTitle border-top border-bottom", 1); 		//Liabilities and owners' equity
	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom", 1); 				//Line no
	tableRow.addCell("期末余额", "valueTitle bold bordersTitle border-top border-bottom", 1); 			//Ending balance
	tableRow.addCell("年初余额", "valueTitle bold border-top border-bottom border-right", 1); 			//At beginning of year

	tableRow = table.addRow();	
	tableRow.addCell("流动资产:", "borders", 4);	 //Current assets
	tableRow.addCell("流动负债:", "borders", 4);	 //Current liabilities

	// 1 - 32
	tableRow = table.addRow();
	tableRow.addCell("货币资金", "borders", 1);
	tableRow.addCell("1", "borders", 1);
	amount = banDoc.currentBalance("Gr=1001|1002|1012", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1001|1002|1012", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("短期借款", "borders", 1);
	tableRow.addCell("32", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2001", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2001", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);
	
	// 2 - 33
	tableRow = table.addRow();
	tableRow.addCell("以公允价值计量且其变动计入当期损益的金融资产", "borders", 1);
	tableRow.addCell("2", "borders", 1);
	amount = banDoc.currentBalance("Gr=1101", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1101", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("以公允价值计量且其变动计入当期损益的金融负债", "borders", 1);
	tableRow.addCell("33", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2101", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2101", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 3 - 34
	tableRow = table.addRow();
	tableRow.addCell("应收票据", "borders", 1);
	tableRow.addCell("3", "borders", 1);
	amount = banDoc.currentBalance("Gr=1121", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1121", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("应付票据", "borders", 1);
	tableRow.addCell("34", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2201", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2201", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 4 : 1122 - 1231
	// 35
	tableRow = table.addRow();
	tableRow.addCell("应收账款", "borders", 1);
	tableRow.addCell("4", "borders", 1);
	amount = banDoc.currentBalance("Gr=1122|1231", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1122|1231", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);

	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("应付账款", "borders", 1);
	tableRow.addCell("35", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2202", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2202", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 5 - 36
	tableRow = table.addRow();
	tableRow.addCell("预付款项", "borders", 1);
	tableRow.addCell("5", "borders", 1);
	amount = banDoc.currentBalance("Gr=1123", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1123", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("预收款项", "borders", 1);
	tableRow.addCell("36", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2203", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2203", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 6 - 37
	tableRow = table.addRow();
	tableRow.addCell("应收利息", "borders", 1);
	tableRow.addCell("6", "borders", 1);
	amount = banDoc.currentBalance("Gr=1132", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1132", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("应付职工薪酬", "borders", 1);
	tableRow.addCell("37", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2211", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2211", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 7 - 38
	tableRow = table.addRow();
	tableRow.addCell("应收股利", "borders", 1);
	tableRow.addCell("7", "borders", 1);
	amount = banDoc.currentBalance("Gr=1131", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1131", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("应交税费", "borders", 1);
	tableRow.addCell("38", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2221", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2221", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 8 - 39
	tableRow = table.addRow();
	tableRow.addCell("其他应收款", "borders", 1);
	tableRow.addCell("8", "borders", 1);
	amount = banDoc.currentBalance("Gr=1221", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1221", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("应付利息", "borders", 1);
	tableRow.addCell("39", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2231", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2231", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 9 - 40
	tableRow = table.addRow();
	tableRow.addCell("存货", "borders", 1);
	tableRow.addCell("9", "borders", 1);
	amount = banDoc.currentBalance("Gr=1321|1401|1402|1403|1404|1405|1406|1407|1408|1411|1461|1471", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1321|1401|1402|1403|1404|1405|1406|1407|1408|1411|1461|1471", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("应付股利", "borders", 1);
	tableRow.addCell("40", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2232", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2232", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	tableRow = table.addRow();
	tableRow.addCell("持有待售资产", "borders", 1);
	tableRow.addCell("", "borders", 1);
	amount = banDoc.currentBalance("Gr=1481|1482", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1481|1482", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount12 = Banana.SDecimal.add(totAmount12,amount);
	totOpening12 = Banana.SDecimal.add(totOpening12,opening);
	tableRow.addCell("其他应付款", "borders", 1);
	tableRow.addCell("", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2241", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2241", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 10 - 41
	tableRow = table.addRow();
	tableRow.addCell("一年内到期的非流动资产", "borders", 1);
	tableRow.addCell("10", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("持有待售负债", "borders", 1);
	tableRow.addCell("41", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2245", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2245", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 11 - 42
	tableRow = table.addRow();
	tableRow.addCell("其他流动资产", "borders", 1);
	tableRow.addCell("11", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("一年内到期的非流动负债", "borders", 1);
	tableRow.addCell("42", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	
	// 12 - 43
	tableRow = table.addRow();
	tableRow.addCell("流动资产合计", "borders styleTotal", 1);
	tableRow.addCell("12", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount12), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening12), "borders right padding-right styleTotal", 1);
	tableRow.addCell("其他流动负债", "borders", 1);
	tableRow.addCell("43", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2401", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2401", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount44 = Banana.SDecimal.add(totAmount44,amount);
	totOpening44 = Banana.SDecimal.add(totOpening44,opening);

	// 44 => total: 32+33+34+35+36+37+38+39+40+xx+41+43
	tableRow = table.addRow();
	tableRow.addCell("", "borders", 1);
	tableRow.addCell("", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("流动负债合计", "borders styleTotal", 1);
	tableRow.addCell("44", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount44), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening44), "borders right padding-right styleTotal", 1);

	tableRow = table.addRow();	
	tableRow.addCell("非流动资产:", "borders", 4);	 //Current assets
	tableRow.addCell("非流动负债:", "borders", 4);	 //Current liabilities

	// 13 - 45
	tableRow = table.addRow();
	tableRow.addCell("可供出售金融资产", "borders", 1);
	tableRow.addCell("13", "borders", 1);
	amount = banDoc.currentBalance("Gr=1503", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1503", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("长期借款", "borders", 1);
	tableRow.addCell("45", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2501", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2501", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount52 = Banana.SDecimal.add(totAmount52,amount);
	totOpening52 = Banana.SDecimal.add(totOpening52,opening);

	// 14 - 46
	tableRow = table.addRow();
	tableRow.addCell("持有至到期投资", "borders", 1);
	tableRow.addCell("14", "borders", 1);
	amount = banDoc.currentBalance("Gr=1501|1502", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1501|1502", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("应付债券", "borders", 1);
	tableRow.addCell("46", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2502", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2502", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount52 = Banana.SDecimal.add(totAmount52,amount);
	totOpening52 = Banana.SDecimal.add(totOpening52,opening);

	// 15 - 47
	tableRow = table.addRow();
	tableRow.addCell("长期应收款", "borders", 1);
	tableRow.addCell("15", "borders", 1);
	amount = banDoc.currentBalance("Gr=1531|1532", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1531|1532", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("长期应付款", "borders", 1);
	tableRow.addCell("47", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2701|2702", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2701|2702", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount52 = Banana.SDecimal.add(totAmount52,amount);
	totOpening52 = Banana.SDecimal.add(totOpening52,opening);

	// 16 - 48
	tableRow = table.addRow();
	tableRow.addCell("长期股权投资", "borders", 1);
	tableRow.addCell("16", "borders", 1);
	amount = banDoc.currentBalance("Gr=1511|1512", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1511|1512", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("专项应付款", "borders", 1);
	tableRow.addCell("48", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2711", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2711", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount52 = Banana.SDecimal.add(totAmount52,amount);
	totOpening52 = Banana.SDecimal.add(totOpening52,opening);

	// 17 - 49
	tableRow = table.addRow();
	tableRow.addCell("投资性房地产", "borders", 1);
	tableRow.addCell("17", "borders", 1);
	amount = banDoc.currentBalance("Gr=1521", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1521", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("预计负债", "borders", 1);
	tableRow.addCell("49", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2801", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2801", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount52 = Banana.SDecimal.add(totAmount52,amount);
	totOpening52 = Banana.SDecimal.add(totOpening52,opening);

	// 18 - 50
	tableRow = table.addRow();
	tableRow.addCell("固定资产", "borders", 1);
	tableRow.addCell("18", "borders", 1);
	amount = banDoc.currentBalance("Gr=1601|1602|1603", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1601|1602|1603", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("递延所得税负债", "borders", 1);
	tableRow.addCell("50", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2901", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=2901", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount52 = Banana.SDecimal.add(totAmount52,amount);
	totOpening52 = Banana.SDecimal.add(totOpening52,opening);

	// 19 - 51
	tableRow = table.addRow();
	tableRow.addCell("工程物资", "borders", 1);
	tableRow.addCell("19", "borders", 1);
	amount = banDoc.currentBalance("Gr=1605", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1605", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("其他非流动负债", "borders", 1);
	tableRow.addCell("51", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);

	// 20
	// 52 => total: 45+46+47+48+49+50
	tableRow = table.addRow();
	tableRow.addCell("在建工程", "borders", 1);
	tableRow.addCell("20", "borders", 1);
	amount = banDoc.currentBalance("Gr=1604", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1604", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("非流动负债合计", "borders styleTotal", 1);
	tableRow.addCell("52", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount52), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening52), "borders right padding-right styleTotal", 1);
	
	// 21
	// 53 => total 44+52
	tableRow = table.addRow();
	tableRow.addCell("固定资产清理", "borders", 1);
	tableRow.addCell("21", "borders", 1);
	amount = banDoc.currentBalance("Gr=1606", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1606", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("负债合计", "borders styleTotal", 1);
	tableRow.addCell("53", "borders styleTotal", 1);
	totAmount53 = Banana.SDecimal.add(totAmount44,totAmount52);
	totOpening53 = Banana.SDecimal.add(totOpening44,totOpening52);
	tableRow.addCell(formatValues(totAmount53), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening53), "borders right padding-right styleTotal", 1);
	
	// 22
	tableRow = table.addRow();
	tableRow.addCell("生产性生物资产", "borders", 1);
	tableRow.addCell("22", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("所有者权益（或股东权益）：", "borders", 1);
	tableRow.addCell("", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	
	// 23
	tableRow = table.addRow();
	tableRow.addCell("油气资产", "borders", 1);
	tableRow.addCell("23", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("实收资本（或股本）", "borders", 1);
	tableRow.addCell("54", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4001", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4001", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount61 = Banana.SDecimal.add(totAmount61,amount);
	totOpening61 = Banana.SDecimal.add(totOpening61,opening);

	// 24 - 55
	tableRow = table.addRow();
	tableRow.addCell("无形资产", "borders", 1);
	tableRow.addCell("24", "borders", 1);
	amount = banDoc.currentBalance("Gr=1701|1702|1703", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1701|1702|1703", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("资本公积", "borders", 1);
	tableRow.addCell("55", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4002", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4002", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount61 = Banana.SDecimal.add(totAmount61,amount);
	totOpening61 = Banana.SDecimal.add(totOpening61,opening);
	
	// 25 - 56
	tableRow = table.addRow();
	tableRow.addCell("开发支出", "borders", 1);
	tableRow.addCell("25", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("减：库存股", "borders", 1);
	tableRow.addCell("56", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4201", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4201", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount61 = Banana.SDecimal.add(totAmount61,amount);
	totOpening61 = Banana.SDecimal.add(totOpening61,opening);
	
	// 26 - 57
	tableRow = table.addRow();
	tableRow.addCell("商誉", "borders", 1);
	tableRow.addCell("26", "borders", 1);
	amount = banDoc.currentBalance("Gr=1711", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1711", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("其他综合收益", "borders", 1);
	tableRow.addCell("57", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	
	// 27 - 58
	tableRow = table.addRow();
	tableRow.addCell("长期待摊费用", "borders", 1);
	tableRow.addCell("27", "borders", 1);
	amount = banDoc.currentBalance("Gr=1801", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1801", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("专项储备", "borders", 1);
	tableRow.addCell("58", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);	

	// 28 - 59
	tableRow = table.addRow();
	tableRow.addCell("递延所得税资产", "borders", 1);
	tableRow.addCell("28", "borders", 1);
	amount = banDoc.currentBalance("Gr=1811", param["startDate"], param["endDate"]).balance;
	opening = banDoc.currentBalance("Gr=1811", param["startDate"], param["endDate"]).opening;
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount30 = Banana.SDecimal.add(totAmount30,amount);
	totOpening30 = Banana.SDecimal.add(totOpening30,opening);
	tableRow.addCell("盈余公积", "borders", 1);
	tableRow.addCell("59", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4101", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4101", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount61 = Banana.SDecimal.add(totAmount61,amount);
	totOpening61 = Banana.SDecimal.add(totOpening61,opening);

	// 29 - 60
	tableRow = table.addRow();
	tableRow.addCell("其他非流动资产", "borders", 1);
	tableRow.addCell("29", "borders", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("", "borders right padding-right", 1);
	tableRow.addCell("未分配利润", "borders", 1);
	tableRow.addCell("60", "borders", 1);
	amount = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4103", param["startDate"], param["endDate"]).balance);
	opening = Banana.SDecimal.invert(banDoc.currentBalance("Gr=4103", param["startDate"], param["endDate"]).opening);
	tableRow.addCell(formatValues(amount), "borders right padding-right", 1);
	tableRow.addCell(formatValues(opening), "borders right padding-right", 1);
	totAmount61 = Banana.SDecimal.add(totAmount61,amount);
	totOpening61 = Banana.SDecimal.add(totOpening61,opening);

	// 30 => total: 13+14+15+16+17+18+19+20+21+24+26+27+28
	// 61 => total: 54+55+56+59+60
	tableRow = table.addRow();
	tableRow.addCell("非流动资产合计", "borders styleTotal", 1);
	tableRow.addCell("30", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount30), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening30), "borders right padding-right styleTotal", 1);
	tableRow.addCell("所有者权益（或股东权益）合计", "borders styleTotal", 1);
	tableRow.addCell("61", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount61), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening61), "borders right padding-right styleTotal", 1);	
	
	// 31 => total: 12+30
	// 62 => total: 53+61
	totAmount31 = Banana.SDecimal.add(totAmount12,totAmount30);
	totOpening31 = Banana.SDecimal.add(totOpening12,totOpening30);
	totAmount62 = Banana.SDecimal.add(totAmount53,totAmount61);
	totOpening62 = Banana.SDecimal.add(totOpening53,totOpening61);
	tableRow = table.addRow();
	tableRow.addCell("资产总计", "borders styleTotal", 1);
	tableRow.addCell("31", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount31), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening31), "borders right padding-right styleTotal", 1);
	tableRow.addCell("负债及所有者权益总计", "borders styleTotal", 1);
	tableRow.addCell("62", "borders styleTotal", 1);
	tableRow.addCell(formatValues(totAmount62), "borders right padding-right styleTotal", 1);
	tableRow.addCell(formatValues(totOpening62), "borders right padding-right styleTotal", 1);	

	//check the balances
	if (Banana.SDecimal.compare(totAmount31,totAmount62) != 0) {
		Banana.document.addMessage("资产负债表不平衡");
	}

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
    pageStyle.setAttribute("margin", "10mm 10mm 5mm 10mm");
    pageStyle.setAttribute("size", "landscape");

    stylesheet.addStyle("body", "font-size: 9pt; font-family: Helvetica");
    stylesheet.addStyle(".bold", "font-weight:bold");
    stylesheet.addStyle(".center", "text-align:center");
    stylesheet.addStyle(".right", "text-align:right");

    stylesheet.addStyle(".border-left", "border-left:thin solid black");
    stylesheet.addStyle(".border-right", "border-right:thin solid black");
    stylesheet.addStyle(".border-top", "border-top:thin solid black");
    stylesheet.addStyle(".border-bottom", "border-bottom:thin solid black");
    stylesheet.addStyle(".padding-right", "padding-right:2px");

	stylesheet.addStyle(".col1", "width:32%");
	stylesheet.addStyle(".col2", "width:%");
	stylesheet.addStyle(".col3", "width:%");
	stylesheet.addStyle(".col4", "width:%");
	stylesheet.addStyle(".col5", "width:32%");
	stylesheet.addStyle(".col6", "width:%");
	stylesheet.addStyle(".col7", "width:%");
	stylesheet.addStyle(".col8", "width:%");

	style = stylesheet.addStyle(".valueTitle");
	style.setAttribute("background-color", "#000000");
	style.setAttribute("color", "#fff");

	style = stylesheet.addStyle(".bordersTitle");
	style.setAttribute("border-right", "thin solid #fff");

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
