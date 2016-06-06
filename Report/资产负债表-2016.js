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
		//"bananaVersion":"Banana Accounting 8", 								//Save the version of Banana Accounting used
		//"scriptVersion":"script v. 2016-04-04",				 				//Save the version of the script
		//"fiscalNumber":banDoc.info("AccountingDataBase","FiscalNumber"),	//Save the fiscal number
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
	form.push({"id":"B1", "gr":"1113", "bClass":"1", "description":"货币资金"});
	form.push({"id":"B2", "gr":"712", "bClass":"1", "description":"短期投资"});
	form.push({"id":"B3", "gr":"113", "bClass":"1", "description":"应收票据"});
	form.push({"id":"B4", "gr":"114", "bClass":"1", "description":"应收账款"});
	form.push({"id":"B5", "gr":"126", "bClass":"1", "description":"预付账款"});
	form.push({"id":"B6", "gr":"712", "bClass":"1", "description":"应收股利"});
	form.push({"id":"B7", "gr":"711", "bClass":"1", "description":"应收利息"});
	form.push({"id":"B8", "gr":"118", "bClass":"1", "description":"其他应收款"});
	form.push({"id":"B9", "gr":"121~122", "bClass":"1", "description":"存货"});
	form.push({"id":"B10", "gr":"121~122", "bClass":"1", "description":"其中:原材料"});
	form.push({"id":"B11", "gr":"121~122", "bClass":"1", "description":"在产品"});
	form.push({"id":"B12", "gr":"121~122", "bClass":"1", "description":"库存商品"});
	form.push({"id":"B13", "gr":"515~518", "bClass":"1", "description":"周转材料"});
	form.push({"id":"B14", "gr":"128~129", "bClass":"1", "description":"其他流动资产"});
	form.push({"id":"B15", "description":"流动资产合计", "sum":"B1;B2;B3;B4;B5;B6;B7;B8;B9;B10;B11;B12;B13;B14"});
	form.push({"id":"B16", "gr":"132", "bClass":"1", "description":"长期债券投资"});
	form.push({"id":"B17", "gr":"132", "bClass":"1", "description":"长期股权投资"});
	form.push({"id":"B18", "gr":"158", "bClass":"1", "description":"固定资产原价"});
	form.push({"id":"B19", "gr":"515~518", "bClass":"1", "description":"减:累计折旧"});
	form.push({"id":"B20", "description":"固定资产账面价值", "sum":"B16;B17;B18;B19"});
	form.push({"id":"B21", "gr":"156", "bClass":"1", "description":"在建工程"});
	form.push({"id":"B22", "gr":"156", "bClass":"1", "description":"工程物资"});
	form.push({"id":"B23", "gr":"158", "bClass":"1", "description":"固定资产清理"});
	form.push({"id":"B24", "gr":"161", "bClass":"1", "description":"生物性生物资产"});
	form.push({"id":"B25", "gr":"17", "bClass":"1", "description":"无形资产"});
	form.push({"id":"B26", "gr":"635~638", "bClass":"1", "description":"开发支出"});
	form.push({"id":"B27", "gr":"515~518", "bClass":"1", "description":"长期待摊费用"});
	form.push({"id":"B28", "gr":"159", "bClass":"1", "description":"其他非流动资产"});
	form.push({"id":"B29", "description":"非流动资产合计", "sum":"B21;B22;B23;B24;B25;B26;B27;B28"});
	form.push({"id":"B30", "description":"资产合计", "sum":"B15;B20;B29"});

	// Liabilities
	form.push({"id":"B31", "gr":"211", "bClass":"2", "description":"短期借款"});
	form.push({"id":"B32", "gr":"213", "bClass":"2", "description":"应付票据"});
	form.push({"id":"B33", "gr":"214", "bClass":"2", "description":"应付账款"});
	form.push({"id":"B34", "gr":"226", "bClass":"2", "description":"预收账款"});
	form.push({"id":"B35", "gr":"217", "bClass":"2", "description":"应付职工薪酬"});
	form.push({"id":"B36", "gr":"217", "bClass":"2", "description":"应交税费"});
	form.push({"id":"B37", "gr":"217", "bClass":"2", "description":"应付利息"});
	form.push({"id":"B38", "gr":"2192", "bClass":"2", "description":"应付利润"});
	form.push({"id":"B39", "gr":"2198", "bClass":"2", "description":"其他应付款"});
	form.push({"id":"B40", "gr":"2194", "bClass":"2", "description":"其他流动负债"});
	form.push({"id":"B41", "description":"流动负债合计", "sum":"B31;B32;B33;B34;B35;B36;B37;B38;B39;B40"});
	form.push({"id":"B42", "gr":"232", "bClass":"2", "description":"长期借款"});
	form.push({"id":"B43", "gr":"2332", "bClass":"2", "description":"长期应付款"});
	form.push({"id":"B44", "gr":"2811", "bClass":"2", "description":"递延收益"});
	form.push({"id":"B45", "gr":"2886", "bClass":"2", "description":"其他非流动负债"});
	form.push({"id":"B46", "description":"非流动负债合计", "sum":"B42;B43;B44;B45"});
	form.push({"id":"B47", "description":"负债合计", "sum":"B41;B46"});
	form.push({"id":"B48", "gr":"3116", "bClass":"2", "description":"实收资本(或股本)"});
	form.push({"id":"B49", "gr":"32", "bClass":"2", "description":"资本公积"});
	form.push({"id":"B50", "gr":"331", "bClass":"2", "description":"盈余公积"});
	form.push({"id":"B51", "gr":"335", "bClass":"2", "description":"未分配利润"});
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
	tableRow.addCell(getValue(form, "B1", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B1", "id").replace("B", ""), "borders", 1);
	tableRow.addCell(getValue(form, "B1", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B1", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B31", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B31", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B31", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B31", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B2", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B2", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B2", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B2", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B32", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B32", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B32", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B32", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B3", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B3", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B3", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B3", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B33", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B33", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B33", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B33", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B4", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B4", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B4", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B4", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B34", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B34", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B34", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B34", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B5", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B5", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B5", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B5", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B35", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B35", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B35", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B35", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B6", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B6", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B6", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B6", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B36", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B36", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B36", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B36", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B7", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B7", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B7", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B7", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B37", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B37", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B37", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B37", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B8", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B8", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B8", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B8", "amount"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B38", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B38", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B38", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B38", "amount"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B9", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B9", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B9", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B9", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B39", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B39", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B39", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B39", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B10", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B10", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B10", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B10", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B40", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B40", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B40", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B40", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B11", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B11", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B11", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B11", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B41", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B41", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B41", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B41", "opening"), "borders right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B12", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B12", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B12", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B12", "opening"), "borders right", 1);
	
	tableRow.addCell("非流动负债:", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B13", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B13", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B13", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B13", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B42", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B42", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B42", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B42", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B14", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B14", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B14", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B14", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B43", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B43", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B43", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B43", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B15", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B15", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B15", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B15", "opening"), "borders right valueTotal", 1);
	
	tableRow.addCell(getValue(form, "B44", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B44", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B44", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B44", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell("非流动资产:", "borders", 4);
	
	tableRow.addCell(getValue(form, "B45", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B45", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B45", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B45", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B16", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B16", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B16", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B16", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B46", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B46", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B46", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B46", "opening"), "borders right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B17", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B17", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B17", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B17", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B47", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B47", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B47", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B47", "opening"), "borders right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B18", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B18", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B18", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B18", "opening"), "borders right", 1);
	
	tableRow.addCell("", "borders", 4);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B19", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B19", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B19", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B19", "opening"), "borders right", 1);
	
	tableRow.addCell("", "borders", 4);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B20", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B20", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B20", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B20", "opening"), "borders right valueTotal", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B21", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B21", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B21", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B21", "opening"), "borders right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B22", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B22", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B22", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B22", "opening"), "borders right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B23", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B23", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B23", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B23", "opening"), "borders right", 1);
	
	tableRow.addCell("", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B24", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B24", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B24", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B24", "opening"), "borders right", 1);
	
	tableRow.addCell("所有者权益(或股东权益):", "borders", 4);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B25", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B25", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B25", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B25", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B48", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B48", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B48", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B48", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B26", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B26", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B26", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B26", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B49", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B49", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B49", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B49", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B27", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B27", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B27", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B27", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B50", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B50", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B50", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B50", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B28", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B28", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B28", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B28", "opening"), "borders right", 1);
	
	tableRow.addCell(getValue(form, "B51", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B51", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B51", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "B51", "opening"), "borders right", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B29", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B29", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B29", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B29", "opening"), "borders right valueTotal", 1);
	
	tableRow.addCell(getValue(form, "B52", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B52", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B52", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B52", "opening"), "borders right valueTotal", 1);



	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "B30", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B30", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B30", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B30", "opening"), "borders right valueTotal", 1);
	
	tableRow.addCell(getValue(form, "B53", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "B53", "id").replace("B", ""), "borders", 1);;
	tableRow.addCell(getValue(form, "B53", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "B53", "opening"), "borders right valueTotal", 1);

	return report;
	/** END PRINT ... */
	/** ------------------------------------------------------------------------------------------------------------- */
}



function loadBalances() {

	for (var i in form) {
	
		if (form[i]["bClass"]) {

			if (form[i]["gr"]) {

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
			}
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














// //Function that create the report
// function createProfitReport(banDoc, startDate, endDate, report) {

// 	/** 1. CREATE AND LOAD THE PARAMETERS AND THE FORM */
// 	loadParam(banDoc, startDate, endDate);
// 	loadForm();
	
// 	/** 2. EXTRACT THE DATA, CALCULATE AND LOAD THE BALANCES */
// 	loadBalances1();
	
// 	/** 3. CALCULATE THE TOTALS */
// 	calcFormTotals(["amount"]);
// 	calcFormTotals(["currentMonthAmount"]);
	
// 	/** 4. DO SOME OPERATIONS BEFORE CONVERTING THE VALUES */
// 	//postProcessAmounts(banDoc);
	
// 	/** 5. CONVERT ALL THE VALUES */
// 	formatValues(["amount"]);
// 	formatValues(["currentMonthAmount"]);
	


// 	/** START PRINT... */
// 	/** ------------------------------------------------------------------------------------------------------------- */

// 	// //Create a report.
// 	// var report = Banana.Report.newReport(param.reportName);

// 	//Table with basic informations
// 	var table = report.addTable("table");

// 	report.addParagraph("利润表(适用执行小企业会计准则的企业)", "bold"); //Balance sheet (application of accounting standards for small business enterprise)
// 	report.addParagraph(" ", "");

// 	//Table with basic informations
// 	var table = report.addTable("table");

// 	var col1 = table.addColumn("col1");
// 	var col2 = table.addColumn("col2");
// 	var col3 = table.addColumn("col3");
// 	var col4 = table.addColumn("col4");
// 	var col5 = table.addColumn("col5");
// 	var col6 = table.addColumn("col6");
// 	var col7 = table.addColumn("col7");
// 	var col8 = table.addColumn("col8");


// 	tableRow = table.addRow();
// 	tableRow.addCell("纳税人基本信息", "center border-bottom", 4); //Taxpayer information
	
// 	tableRow = table.addRow();
// 	tableRow.addCell("纳税人识别号:", "", 1); //Taxpayer identification number:
// 	tableRow.addCell("xxxx", "", 1);

// 	tableRow.addCell("纳税人名称:", "", 1); //Payer's name
// 	tableRow.addCell("xxxx", "", 1);

// 	tableRow = table.addRow();
// 	var d = new Date();
// 	tableRow.addCell("报送日期: " + Banana.Converter.toLocaleDateFormat(d), "", 2); //Submit date
// 	//tableRow.addCell(, "", 3);

// 	tableRow.addCell("所属时期起: " + Banana.Converter.toLocaleDateFormat(startDate), "", 1); //Period

// 	tableRow.addCell("所属时期止: " + Banana.Converter.toLocaleDateFormat(endDate), "", 1); //Period

// 	tableRow = table.addRow();
// 	tableRow.addCell("", "", 4);


// 	//----------------------------------------------------------------------------------------------//
// 	// //Table with basic informations
// 	// var table1 = report.addTable("table1");	
// 	// var tableHeader = table1.getHeader();
// 	tableRow = table.addRow();	
// 	tableRow.addCell("项   目", "valueTitle bold bordersTitle border-left border-top border-bottom center", 1); //Project
// 	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom center", 1); 				//Line no
// 	tableRow.addCell("本年累计金额", "valueTitle bold bordersTitle border-top border-bottom center", 1); 		//YTD amount
// 	tableRow.addCell("本月金额", "valueTitle bold bordersTitle border-top border-bottom center", 1); 			//this month amount




// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P1", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P1", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P1", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P1", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P2", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P2", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P2", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P2", "currentMonthAmount"), "borders right", 1);

// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P3", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P3", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P3", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P3", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P4", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P4", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P4", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P4", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P5", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P5", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P5", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P5", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P6", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P6", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P6", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P6", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P7", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P7", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P7", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P7", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P8", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P8", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P8", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P8", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P9", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P9", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P9", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P9", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P10", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P10", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P10", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P10", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P11", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P11", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P11", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P11", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P12", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P12", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P12", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P12", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P13", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P13", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P13", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P13", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P14", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P14", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P14", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P14", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P15", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P15", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P15", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P15", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P16", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P16", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P16", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P16", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P17", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P17", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P17", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P17", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P18", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P18", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P18", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P18", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P19", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P19", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P19", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P19", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P20", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P20", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P20", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P20", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P21", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P21", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P21", "amount"), "borders right valueTotal", 1);
// 	tableRow.addCell(getValue(form, "P21", "currentMonthAmount"), "borders right valueTotal", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P22", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P22", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P22", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P22", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P23", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P23", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P23", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P23", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P24", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P24", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P24", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P24", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P25", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P25", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P25", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P25", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P26", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P26", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P26", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P26", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P27", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P27", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P27", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P27", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P28", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P28", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P28", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P28", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P29", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P29", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P29", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P29", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P30", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P30", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P30", "amount"), "borders right valueTotal", 1);
// 	tableRow.addCell(getValue(form, "P30", "currentMonthAmount"), "borders right valueTotal", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P31", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P31", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P31", "amount"), "borders right", 1);
// 	tableRow.addCell(getValue(form, "P31", "currentMonthAmount"), "borders right", 1);


// 	tableRow = table.addRow();
// 	tableRow.addCell(getValue(form, "P32", "description"), "borders", 1);
// 	tableRow.addCell(getValue(form, "P32", "id").replace("P", ""), "borders", 1);
// 	tableRow.addCell(getValue(form, "P32", "amount"), "borders right valueTotal", 1);
// 	tableRow.addCell(getValue(form, "P32", "currentMonthAmount"), "borders right valueTotal", 1);


// 	//Add a footer to the report
// 	//addFooter(report, param);


// 	/** END PRINT ... */
// 	/** ------------------------------------------------------------------------------------------------------------- */

// 	//Return the final report
// 	//return report;
// }







/*

//The purpose of this function is to load all the balances and save the values into the form
function loadBalances() {

	for (var i in form) {

		//Check if there are "vatClass" properties, then load VAT balances
		if (form[i]["vatClass"]) {
			if (form[i]["gr"]) {
				form[i]["amount"] = calculateVatGr1Balance(form[i]["gr"], form[i]["vatClass"], param["grColumn"], param["startDate"], param["endDate"]);
			}
		}

		//Check if there are "bClass" properties, then load balances
		if (form[i]["bClass"]) {
			if (form[i]["gr"]) {
				form[i]["amount"] = calculateAccountGr1Balance(form[i]["gr"], form[i]["bClass"], param["grColumn"], param["startDate"], param["endDate"]);
				form[i]["opening"] = calculateAccountGr1Opening(form[i]["gr"], form[i]["bClass"], param["grColumn"], param["startDate"], param["endDate"]);
			}
		}
	}
}


//The purpose of this function is to calculate all the balances of the accounts belonging to the same group (grText)
function calculateAccountGr1Balance(grText, bClass, grColumn, startDate, endDate) {
	
	var accounts = getColumnListForGr(Banana.document.table("Accounts"), grText, "Account", grColumn);	
	accounts = accounts.join("|");
	
	//Sum the amounts of opening, debit, credit, total and balance for all transactions for this accounts
	var currentBal = Banana.document.currentBalance(accounts, startDate, endDate);
	
	//The "bClass" decides which value to use
	if (bClass === "0") {
		return currentBal.amount;
	}
	else if (bClass === "1") {
		return currentBal.balance;
	}
	else if (bClass === "2") {
		return Banana.SDecimal.invert(currentBal.balance);
	}
	else if (bClass === "3") {
		return currentBal.total;
	}
	else if (bClass === "4") {
		return Banana.SDecimal.invert(currentBal.total);
	}
}



//The purpose of this function is to calculate all the balances of the accounts belonging to the same group (grText)
function calculateAccountGr1Opening(grText, bClass, grColumn, startDate, endDate) {
	
	var accounts = getColumnListForGr(Banana.document.table("Accounts"), grText, "Account", grColumn);	
	accounts = accounts.join("|");
	
	//Sum the amounts of opening, debit, credit, total and balance for all transactions for this accounts
	var currentBal = Banana.document.currentBalance(accounts, startDate, endDate);
	
	//The "bClass" decides which value to use
	// if (bClass === "0") {
	// 	return currentBal.amount;
	// }
	//else 
	if (bClass === "1") {
		return currentBal.opening;
	}
	else if (bClass === "2") {
		return Banana.SDecimal.invert(currentBal.opening);
	}
	// else if (bClass === "3") {
	// 	return currentBal.total;
	// }
	// else if (bClass === "4") {
	// 	return Banana.SDecimal.invert(currentBal.total);
	// }
}



//The main purpose of this function is to create an array with all the values of a given column of the table (codeColumn) belonging to the same group (grText)
function getColumnListForGr(table, grText, codeColumn, grColumn) {

	if (table === undefined || !table) {
		return str;
	}

	if (!grColumn) {
		grColumn = "Gr1";
	}

	var str = [];

	//Loop to take the values of each rows of the table
	for (var i = 0; i < table.rowCount; i++) {
		var tRow = table.row(i);
		var grRow = tRow.value(grColumn);

		//If Gr1 column contains other characters (in this case ";") we know there are more values
		//We have to split them and take all values separately
		//If there are only alphanumeric characters in Gr1 column we know there is only one value
		var codeString = grRow;
		var arrCodeString = codeString.split(";");
		for (var j = 0; j < arrCodeString.length; j++) {
			var codeString1 = arrCodeString[j];
			if (codeString1 === grText) {
				str.push(tRow.value(codeColumn));
			}
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

	//Return the array
	return str;
}

*/



// //The purpose of this function is to calculate all the VAT balances of the accounts belonging to the same group (grText)
// function calculateVatGr1Balance(grText, vatClass, grColumn, startDate, endDate) {
	
// 	var grCodes = getColumnListForGr(Banana.document.table("VatCodes"), grText, "VatCode", grColumn);
// 	grCodes = grCodes.join("|");

// 	//Sum the vat amounts for the specified vat code and period
// 	var currentBal = Banana.document.vatCurrentBalance(grCodes, startDate, endDate);

// 	//The "vatClass" decides which value to use
// 	if (vatClass === "1") {
// 		if (currentBal.vatTaxable != 0) {
// 			return currentBal.vatTaxable;
// 		}
// 	}
// 	else if (vatClass === "2") {
// 		if (currentBal.vatTaxable != 0) {
// 			return Banana.SDecimal.invert(currentBal.vatTaxable);
// 		}
// 	}
// 	else if (vatClass === "3") {
// 		if (currentBal.vatPosted != 0) {
// 			return currentBal.vatPosted;
// 		}
// 	}
// 	else if (vatClass === "4") {
// 		if (currentBal.vatPosted != 0) {
// 			return Banana.SDecimal.invert(currentBal.vatPosted);
// 		}
// 	}
// }


// //This function return the difference in months between two dates
// function getMonthDiff(d1, d2) {	
// 	var months;
//     months = (d2.getFullYear() - d1.getFullYear()) * 12;
//     //months -= d1.getMonth() + 1;
// 	months -= d1.getMonth();
//     months += d2.getMonth();
//     //Increment months if d2 comes later in its month than d1 in its month
//     if (d2.getDate() >= d1.getDate()) {
//         months++;
//     }
// 	return months <= 0 ? 0 : months;
// }


// //Function to get the name of the months (German)
// function getMonthName(date) {
// 	var month = [];
// 	month[0] = "Januar";
// 	month[1] = "Februar";
// 	month[2] = "März";
// 	month[3] = "April";
// 	month[4] = "Mai";
// 	month[5] = "Juni";
// 	month[6] = "Juli";
// 	month[7] = "August";
// 	month[8] = "September";
// 	month[9] = "Oktober";
// 	month[10] = "November";
// 	month[11] = "Dezember";
// 	var monthName = month[date.getMonth()];
// 	return monthName;
// }



// //This function adds a Footer to the report
// function addFooter(report, param) {
//    report.getFooter().addClass("footer");
//    var versionLine = report.getFooter().addText(param.bananaVersion + ", " + param.scriptVersion + ", ", "description");
//    //versionLine.excludeFromTest();
//    report.getFooter().addText(param.pageCounterText + " ", "description");
//    report.getFooter().addFieldPageNr();
// }


// //The purpose of this function is to verify two sums.
// //Given two lists of values divided by the character ";" the function creates two totals and compares them.
// //It is also possible to compare directly single values, instead of a list of values.
// function checkTotals(valuesList1, valuesList2) {
// 	//Calculate the first total
// 	if (valuesList1) {
// 		var total1 = 0;
// 		var arr1 = valuesList1.split(";");
// 		for (var i = 0; i < arr1.length; i++) {
// 			total1 = Banana.SDecimal.add(total1, getValue(form, arr1[i], "amount"), {'decimals':param.rounding});
// 		}
// 	}
	
// 	//Calculate the second total
// 	if (valuesList2) {
// 		var total2 = 0;
// 		var arr2 = valuesList2.split(";");
// 		for (var i = 0; i < arr2.length; i++) {
// 			total2 = Banana.SDecimal.add(total2, getValue(form, arr2[i], "amount"), {'decimals':param.rounding});
// 		}
// 	}
	
// 	//Finally, compare the two totals.
// 	//If there are differences, a message and a dialog box warns the user
// 	if (Banana.SDecimal.compare(total1, total2) !== 0) {
		
// 		//Add an information dialog.
// 		Banana.Ui.showInformation("Warning!", "Different values: Total " + valuesList1 + " <" + Banana.Converter.toLocaleNumberFormat(total1) + 
// 		">, Total " + valuesList2 + " <" + Banana.Converter.toLocaleNumberFormat(total2) + ">");
		
// 		//Add to the form an object containing a warning message that will be added at the end of the report
// 		var warningStringMsg = "Warning! Different values: Total " + valuesList1 + " <" + Banana.Converter.toLocaleNumberFormat(total1) + 
// 							">, Total " + valuesList2 + " <" + Banana.Converter.toLocaleNumberFormat(total2) + ">";
		
// 		form.push({"warningMessage" : warningStringMsg});
// 	}
// }


// //The purpose of this function is to verify if the balance from Banana euquals the report total
// function checkBalance(banDoc) {
// 	//First, we get the total from the report, specifying the correct id total 
// 	var totalFromReport = getValue(form, "7", "amount");

// 	//Second, we get the VAT balance table from Banana using the function Banana.document.vatReport([startDate, endDate]).
// 	//The two dates are taken directly from the structure. 
// 	var vatReportTable = banDoc.vatReport(param.startDate, param.endDate);
	
// 	//Now we can read the table rows values
// 	for (var i = 0; i < vatReportTable.rowCount; i++) {
// 		var tRow = vatReportTable.row(i);
// 		var group = tRow.value("Group");
// 		var vatBalance = tRow.value("VatBalance");
		
// 		//Since we know that the balance is summed in group named "_tot_", we check if that value equals the total from the report
// 		if (group === "_tot_") {

// 			//In order to compare correctly the values we have to invert the sign of the result from Banana (if negative), using the Banana.SDecimal.invert() function.
// 			if (Banana.SDecimal.sign(vatBalance) == -1) {
// 				vatBalance = Banana.SDecimal.invert(vatBalance);
// 			}

// 			//Now we can compare the two values using the Banana.SDecimal.compare() function and return a message if they are different.
// 			if (Banana.SDecimal.compare(totalFromReport, vatBalance) !== 0) {

// 				//Add an information dialog
// 				Banana.Ui.showInformation("Warning!", "Different values: " + 
// 				"Total from Banana <" + Banana.Converter.toLocaleNumberFormat(vatBalance) + 
// 				">, Total from report <" + Banana.Converter.toLocaleNumberFormat(totalFromReport) + ">");

// 				//Add to the form an object containing a warning message that will be added at the end of the report
// 				var warningStringMsg =	"Warning! Different values: " + "Total from Banana <" + Banana.Converter.toLocaleNumberFormat(vatBalance) + 
// 										">, Total from report <" + Banana.Converter.toLocaleNumberFormat(totalFromReport) + ">";
				
// 				form.push({"warningMessage" : warningStringMsg});
// 			}
// 		}
// 	}
// }