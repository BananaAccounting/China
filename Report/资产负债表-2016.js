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
// @id = ch.banana.addon.balancereportchina2016
// @api = 1.0
// @pubdate = 2016-04-04
// @publisher = Banana.ch SA
// @description = 资产负债表 2016
// @task = app.command
// @doctype = 100.*
// @docproperties = china
// @outputformat = none
// @inputdataform = none
// @timeout = -1




//Global variables
var form = [];
var param = {};



function loadParam(banDoc, startDate, endDate) {
	param = {
		"reportName":"China - Balance Report 2016",							//Save the report's name
		"startDate":startDate,												//Save the startDate that will be used to specify the accounting period starting date
		"endDate":endDate, 													//Save the endDate that will be used to specify the accounting period ending date		
		"taxpayerNumber":banDoc.info("AccountingDataBase","FiscalNumber"),
		"company":banDoc.info("AccountingDataBase","Company"), 				//Save the company name
		"address":banDoc.info("AccountingDataBase","Address1"), 			//Save the address																	//Save details
		"nation":banDoc.info("AccountingDataBase","Country"), 				//Save the country
		"telephone":banDoc.info("AccountingDataBase","Phone"), 				//Save the phone number
		"zip":banDoc.info("AccountingDataBase","Zip"), 						//Save the zip code
		"city":banDoc.info("AccountingDataBase","City"),					//Save the city
		"pageCounterText":"",												//Save the text for the page counter
		"grColumn" : "Gr",													//Save the GR column (Gr1 or Gr2)
		"rounding" : 2,														//Speficy the rounding type		
		"formatNumber":true 												//Choose if format number or not
	};
}


function loadForm() {

	/*
		BALANCE
	*/

	//Assets
	form.push({"id":"B1", "gr":"100", "bClass":"1", "description":"货币资金"});
	form.push({"id":"B2", "gr":"1101", "bClass":"1", "description":"短期投资"});
	form.push({"id":"B3", "gr":"1201", "bClass":"1", "description":"应收票据"});
	form.push({"id":"B4", "gr":"1202", "bClass":"1", "description":"应收账款"});
	form.push({"id":"B5", "gr":"1203", "bClass":"1", "description":"预付账款"});
	form.push({"id":"B6", "gr":"1131", "bClass":"1", "description":"应收股利"});
	form.push({"id":"B7", "gr":"1221", "bClass":"1", "description":"应收利息"});
	form.push({"id":"B8", "gr":"1231", "bClass":"1", "description":"其他应收款"});
	form.push({"id":"B9", "gr":"130", "bClass":"1", "description":"存货"});
	form.push({"id":"B10", "gr":"1302", "bClass":"1", "description":"其中:原材料"});
	form.push({"id":"B11", "gr":"1301", "bClass":"1", "description":"在产品"});
	form.push({"id":"B12", "gr":"1303", "bClass":"1", "description":"库存商品"});
	form.push({"id":"B13", "gr":"1431", "bClass":"1", "description":"周转材料"});
	form.push({"id":"B14", "gr":"", "bClass":"1", "description":"其他流动资产"});
	form.push({"id":"B15", "description":"流动资产合计", "sum":"B1;B2;B3;B4;B5;B6;B7;B8;B9;B10;B11;B12;B13;B14"});
	form.push({"id":"B16", "gr":"140", "bClass":"1", "description":"长期债券投资"});
	form.push({"id":"B17", "gr":"141", "bClass":"1", "description":"长期股权投资"});
	form.push({"id":"B18", "gr":"1501", "bClass":"1", "description":"固定资产原价"});
	form.push({"id":"B19", "gr":"1502", "bClass":"1", "description":"减:累计折旧"});
	form.push({"id":"B20", "description":"固定资产账面价值", "sum":"B18;-B19"});
	form.push({"id":"B21", "gr":"1503", "bClass":"1", "description":"在建工程"});
	form.push({"id":"B22", "gr":"1605", "bClass":"1", "description":"工程物资"});
	form.push({"id":"B23", "gr":"1504", "bClass":"1", "description":"固定资产清理"});
	form.push({"id":"B24", "gr":"1511", "bClass":"1", "description":"生物性生物资产"});
	form.push({"id":"B25", "gr":"160", "bClass":"1", "description":"无形资产"});
	form.push({"id":"B26", "gr":"", "bClass":"1", "description":"开发支出"});
	form.push({"id":"B27", "gr":"170", "bClass":"1", "description":"长期待摊费用"});
	form.push({"id":"B28", "gr":"", "bClass":"1", "description":"其他非流动资产"});
	form.push({"id":"B29", "description":"非流动资产合计", "sum":"B20;B21;B22;B23;B24;B25;B26;B27;B28"});
	form.push({"id":"B30", "description":"资产合计", "sum":"B15;B29"});

	// Liabilities
	form.push({"id":"B31", "gr":"200", "bClass":"2", "description":"短期借款"});
	form.push({"id":"B32", "gr":"", "bClass":"2", "description":"应付票据"});
	form.push({"id":"B33", "gr":"210", "bClass":"2", "description":"应付账款"});
	form.push({"id":"B34", "gr":"211", "bClass":"2", "description":"预收账款"});
	form.push({"id":"B35", "gr":"220", "bClass":"2", "description":"应付职工薪酬"});
	form.push({"id":"B36", "gr":"230", "bClass":"2", "description":"应交税费"});
	form.push({"id":"B37", "gr":"240", "bClass":"2", "description":"应付利息"});
	form.push({"id":"B38", "gr":"250", "bClass":"2", "description":"应付利润"});
	form.push({"id":"B39", "gr":"260", "bClass":"2", "description":"其他应付款"});
	form.push({"id":"B40", "gr":"", "bClass":"2", "description":"其他流动负债"});
	form.push({"id":"B41", "description":"流动负债合计", "sum":"B31;B32;B33;B34;B35;B36;B37;B38;B39;B40"});
	form.push({"id":"B42", "gr":"270", "bClass":"2", "description":"长期借款"});
	form.push({"id":"B43", "gr":"", "bClass":"2", "description":"长期应付款"});
	form.push({"id":"B44", "gr":"280", "bClass":"2", "description":"递延收益"});
	form.push({"id":"B45", "gr":"", "bClass":"2", "description":"其他非流动负债"});
	form.push({"id":"B46", "description":"非流动负债合计", "sum":"B42;B43;B44;B45"});
	form.push({"id":"B47", "description":"负债合计", "sum":"B41;B46"});
	form.push({"id":"B48", "gr":"400", "bClass":"2", "description":"实收资本(或股本)"});
	form.push({"id":"B49", "gr":"401", "bClass":"2", "description":"资本公积"});
	form.push({"id":"B50", "gr":"410", "bClass":"2", "description":"盈余公积"});
	form.push({"id":"B51", "gr":"410401", "bClass":"2", "description":"未分配利润"});
	form.push({"id":"B52", "description":"所有者权益(或股东权益)合计", "sum":"B48;B49;B50;B51"});
	form.push({"id":"B53", "description":"负债和所有者权益(或股东权益)总计", "sum":"B47;B52"});

}



//------------------------------------------------------------------------------//
// FUNCTIONS
//------------------------------------------------------------------------------//

//Main function
function exec(string) {
	//Check if we are on an opened document
	if (!Banana.document) {
		return;
	}

	//Open a dialog window asking to select an item from the list
	//var itemSelected = Banana.Ui.getItem('Input', 'Choose a value', ["1. 资产负债表","2. 损益表"], 0, false);

	//Function call to manage and save user settings about the period date
	var dateform = getPeriodSettings();

	//Check if user has entered a period date.
	//If yes, we can get all the informations we need, process them and finally create the report.
	//If not, the script execution will be stopped immediately.
	if (dateform) {

		//Function call to create the report
		var report = createBalanceReport(Banana.document, dateform.selectionStartDate, dateform.selectionEndDate);

		//Print the report
		var stylesheet = createStylesheet();
		Banana.Report.preview(report, stylesheet);	
	}
}


//The purpose of this function is to do some operations before the values are converted
function postProcessAmounts(banDoc) {}


//Function that create the report
function createBalanceReport(banDoc, startDate, endDate, report) {

	/** 1. CREATE AND LOAD THE PARAMETERS AND THE FORM */
	loadParam(banDoc, startDate, endDate);
	loadForm();
	
	/** 2. EXTRACT THE DATA, CALCULATE AND LOAD THE BALANCES */
	loadBalances()
	
	/** 3. CALCULATE THE TOTALS */
	calcFormTotals(["amount"]);
	calcFormTotals(["opening"]);
	
	/** 4. DO SOME OPERATIONS BEFORE CONVERTING THE VALUES */
	//postProcessAmounts(banDoc);
	
	/** 5. CONVERT ALL THE VALUES */
	formatValues(["amount"]);
	formatValues(["opening"]);


	/** START PRINT... */
	/** ------------------------------------------------------------------------------------------------------------- */

	//Create a report.
	var report = Banana.Report.newReport(param.reportName);

	//Table with basic informations
	var table = report.addTable("table");

	report.addParagraph("资产负债表(适用执行小企业会计准则的企业) - (ALFA VERSION)", "bold"); //Balance sheet (application of accounting standards for small business enterprise)

	//Table with basic informations
	var table = report.addTable("table");

	var col1 = table.addColumn("col1");
	var col2 = table.addColumn("col2");
	var col3 = table.addColumn("col3");
	var col4 = table.addColumn("col4");
	var col5 = table.addColumn("col5");
	var col6 = table.addColumn("col6");
	var col7 = table.addColumn("col7");
	var col8 = table.addColumn("col8");


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


	//----------------------------------------------------------------------------------------------//
	// var table1 = report.addTable("table1");	
	// var tableHeader = table1.getHeader();
	tableRow = table.addRow();	
	tableRow.addCell("资 产", "valueTitle bold bordersTitle border-left border-top border-bottom", 1); 	//Assets
	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom", 1); 				//Line no
	tableRow.addCell("期末余额", "valueTitle bold bordersTitle border-top border-bottom", 1); 			//Ending balance
	tableRow.addCell("年初余额", "valueTitle bold bordersTitle border-top border-bottom", 1); 			//At beginning of year
	tableRow.addCell("负债和所有者权益", "valueTitle bold bordersTitle border-top border-bottom", 1); 		//Liabilities and owners' equity
	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom", 1); 				//Line no
	tableRow.addCell("期末余额", "valueTitle bold bordersTitle border-top border-bottom", 1); 			//Ending balance
	tableRow.addCell("年初余额", "valueTitle bold border-top border-bottom border-right", 1); 			//At beginning of year

	tableRow = table.addRow();	
	tableRow.addCell("流动资产:", "borders", 4);	 //Current assets
	tableRow.addCell("流动负债:", "borders", 4);	 //Current liabilities




	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B1", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B1", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B1", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B1", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B31", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B31", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B31", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B31", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B2", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B2", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B2", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B2", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B32", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B32", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B32", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B32", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B3", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B3", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B3", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B3", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B33", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B33", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B33", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B33", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B4", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B4", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B4", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B4", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B34", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B34", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B34", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B34", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B5", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B5", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B5", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B5", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B35", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B35", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B35", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B35", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B6", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B6", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B6", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B6", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B36", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B36", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B36", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B36", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B7", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B7", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B7", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B7", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B37", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B37", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B37", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B37", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B8", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B8", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B8", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B8", "amount"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B38", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B38", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B38", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B38", "amount"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B9", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B9", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B9", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B9", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B39", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B39", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B39", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B39", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B10", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B10", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B10", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B10", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B40", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B40", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B40", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B40", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B11", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B11", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B11", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B11", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B41", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B41", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B41", "amount"), "borders right  padding-right valueTotal", 1);
	tableRow.addCell(getValue(form, "B41", "opening"), "borders right  padding-right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B12", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B12", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B12", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B12", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("非流动负债:", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B13", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B13", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B13", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B13", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B42", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B42", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B42", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B42", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B14", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B14", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B14", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B14", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B43", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B43", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B43", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B43", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B15", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B15", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B15", "amount"), "borders right  padding-right valueTotal", 1);
	tableRow.addCell(getValue(form, "B15", "opening"), "borders right  padding-right valueTotal", 1);
	
	tableRow.addCell(getValue(form, "B44", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B44", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B44", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B44", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell("非流动资产:", "borders", 4);
	
	tableRow.addCell(getValue(form, "B45", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B45", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B45", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B45", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B16", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B16", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B16", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B16", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B46", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B46", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B46", "amount"), "borders right padding-right valueTotal", 1);
	tableRow.addCell(getValue(form, "B46", "opening"), "borders right padding-right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B17", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B17", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B17", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B17", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B47", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B47", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B47", "amount"), "borders right  padding-right valueTotal", 1);
	tableRow.addCell(getValue(form, "B47", "opening"), "borders right  padding-right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B18", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B18", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B18", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B18", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("", "borders", 4);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B19", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B19", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B19", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B19", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("", "borders", 4);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B20", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B20", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B20", "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, "B20", "opening"), "borders right valueTotal padding-right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B21", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B21", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B21", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B21", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B22", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B22", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B22", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B22", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B23", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B23", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B23", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B23", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B24", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B24", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B24", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B24", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell("所有者权益(或股东权益):", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B25", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B25", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B25", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B25", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B48", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B48", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B48", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B48", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B26", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B26", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B26", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B26", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B49", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B49", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B49", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B49", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B27", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B27", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B27", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B27", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B50", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B50", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B50", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B50", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B28", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B28", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B28", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B28", "opening"), "borders right padding-right", 1);
	
	tableRow.addCell(getValue(form, "B51", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B51", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B51", "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, "B51", "opening"), "borders right padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B29", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B29", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B29", "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, "B29", "opening"), "borders right valueTotal padding-right", 1);
	
	tableRow.addCell(getValue(form, "B52", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B52", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B52", "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, "B52", "opening"), "borders right valueTotal padding-right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B30", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B30", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B30", "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, "B30", "opening"), "borders right valueTotal padding-right", 1);
	
	tableRow.addCell(getValue(form, "B53", "description"), "borders padding-left", 1);
	tableRow.addCell(getValue(form, "B53", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B53", "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, "B53", "opening"), "borders right valueTotal padding-right", 1);

	return report;
	/** END PRINT ... */
	/** ------------------------------------------------------------------------------------------------------------- */
}



function loadBalances() {

	for (var i in form) {
	
		if (form[i]["bClass"]) {

			//if (form[i]["gr"]) {

				var bClass = form[i]["bClass"];

				//Sum the amounts of opening, debit, credit, total and balance for all transactions for this accounts
				// gr = "Gr=1113|111"
				var currentBal = Banana.document.currentBalance("Gr="+form[i]["gr"], param["startDate"], param["endDate"]);
				
				form[i]["amount"] = "";
				form[i]["opening"] = "";

				//The "bClass" decides which value to use
				if (bClass === "0") {
					form[i]["amount"] = currentBal.balance;
				}
				else if (bClass === "1") {
					form[i]["amount"]  = currentBal.balance;
					form[i]["opening"]  = currentBal.opening;
				}
				else if (bClass === "2") {
					form[i]["amount"]  = Banana.SDecimal.invert(currentBal.balance);
					form[i]["opening"]  = Banana.SDecimal.invert(currentBal.opening);
				}
				else if (bClass === "3") {
					form[i]["amount"]  = currentBal.total;
				}
				else if (bClass === "4") {
					form[i]["amount"]  = Banana.SDecimal.invert(currentBal.total);
				}
			//}
		}
	}
}


//The purpose of this function is to convert all the values from the given list to local format
function formatValues(fields) {
	if (param["formatNumber"] === true) {
		for (i = 0; i < form.length; i++) {
			var valueObj = getObject(form, form[i].id);

			for (var j = 0; j < fields.length; j++) {
				valueObj[fields[j]] = Banana.Converter.toLocaleNumberFormat(valueObj[fields[j]]);
			}
		}
	}
}


//Calculate all totals of the form
function calcFormTotals(fields) {
	for (var i = 0; i < form.length; i++) {
		calcTotal(form[i].id, fields);
	}
}


//Calculate a total of the form
function calcTotal(id, fields) {
	
	var valueObj = getObject(form, id);
	
	if (valueObj[fields[0]]) { //first field is present
		return; //calc already done, return
	}
	
	if (valueObj.sum) {
		var sumElements = valueObj.sum.split(";");	
		
		for (var k = 0; k < sumElements.length; k++) {
			var entry = sumElements[k].trim();
			if (entry.length <= 0) {
				return true;
			}
			
			var isNegative = false;
			if (entry.indexOf("-") >= 0) {
				isNegative = true;
				entry = entry.substring(1);
			}
			
			//Calulate recursively
			calcTotal(entry, fields);  
			
		    for (var j = 0; j < fields.length; j++) {
				var fieldName = fields[j];
				var fieldValue = getValue(form, entry, fieldName);
				if (fieldValue) {
					if (isNegative) {
						//Invert sign
						fieldValue = Banana.SDecimal.invert(fieldValue);
					}
					valueObj[fieldName] = Banana.SDecimal.add(valueObj[fieldName], fieldValue, {'decimals':param.rounding});
				}
			}
		}
	} else if (valueObj.gr) {
		//Already calculated in loadFormBalances()
	}
}


//The purpose of this function is to return a specific field value from the object.
//When calling this function, it's necessary to speficy the form (the structure), the object ID, and the field (parameter) needed.
function getValue(form, id, field) {
	var searchId = id.trim();
	for (var i = 0; i < form.length; i++) {
		if (form[i].id === searchId) {
			return form[i][field];
		}
	}
	Banana.document.addMessage("Couldn't find object with id:" + id);
}


//This function is very similar to the getValue() function.
//Instead of returning a specific field from an object, this function return the whole object.
function getObject(form, id) {
	for (var i = 0; i < form.length; i++) {
		if (form[i]["id"] === id) {
			return form[i];
		}
	}
	Banana.document.addMessage("Couldn't find object with id: " + id);
}


//The main purpose of this function is to allow the user to enter the accounting period desired and saving it for the next time the script is run.
//Every time the user runs of the script he has the possibility to change the date of the accounting period.
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


//The main purpose of this function is to create styles for the report print
function createStylesheet() {
	var stylesheet = Banana.Report.newStyleSheet();

    var pageStyle = stylesheet.addStyle("@page");
    pageStyle.setAttribute("margin", "10mm 10mm 5mm 10mm");
    pageStyle.setAttribute("size", "landscape");

    stylesheet.addStyle("body", "font-size: 10pt; font-family: Helvetica");
    stylesheet.addStyle(".bold", "font-weight:bold");
    stylesheet.addStyle(".center", "text-align:center");
    stylesheet.addStyle(".right", "text-align:right");

    stylesheet.addStyle(".border-left", "border-left:thin solid black");
    stylesheet.addStyle(".border-right", "border-right:thin solid black");
    stylesheet.addStyle(".border-top", "border-top:thin solid black");
    stylesheet.addStyle(".border-bottom", "border-bottom:thin solid black");

    stylesheet.addStyle(".padding-left", "padding-left:12px");
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
	//style.setAttribute("padding-bottom", "5px"); 
	style.setAttribute("background-color", "#000000");
	style.setAttribute("color", "#fff");

	style = stylesheet.addStyle(".bordersTitle");
	style.setAttribute("border-right", "thin solid #fff");
	//style.setAttribute("padding-bottom", "2px");
	//style.setAttribute("padding-top", "5px"); 

	style = stylesheet.addStyle(".borders");
	style.setAttribute("border-left", "thin solid black");
	style.setAttribute("border-right", "thin solid black");
	style.setAttribute("border-bottom", "thin solid black");
	style.setAttribute("border-top", "thin solid black");
	//style.setAttribute("padding-bottom", "2px");
	//style.setAttribute("padding-top", "5px"); 

	style = stylesheet.addStyle(".valueTotal");
	style.setAttribute("font-weight", "bold");
	style.setAttribute("background-color", "#eeeeee");

	style = stylesheet.addStyle(".footer");
	style.setAttribute("text-align", "right");
	style.setAttribute("font-size", "8px");
	style.setAttribute("font-family", "Courier New");

	style = stylesheet.addStyle(".table");
	style.setAttribute("width", "100%");
	//stylesheet.addStyle("table.table td", "border: thin solid black");

	return stylesheet;
}


