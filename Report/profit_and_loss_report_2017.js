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
// @id = ch.banana.addon.china.report.profitloss2017
// @api = 1.0
// @pubdate = 2017-04-24
// @publisher = Banana.ch SA
// @description.zh = 损益表
// @description.en = Profit and Loss 2017
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @docproperties = china
// @outputformat = none
// @inputdataform = none
// @timeout = -1




//Global variables
var form = [];
var param = {};
/* the profit and loss Gr1 in Chinese */
var GR = {};
	


function loadParam(banDoc, startDate, endDate) {
	var date = new Date();
	param = {
		"reportName":"China - Profit/Loss Report 2017",						//Save the report's name
		"startDate":startDate,												//Save the startDate that will be used to specify the accounting period starting date
		"endDate":endDate,													//Save the endDate that will be used to specify the accounting period ending date
		"currentDate": date,												//Save the current Date
		"taxpayerNumber":banDoc.info("AccountingDataBase","FiscalNumber"),
		"company":banDoc.info("AccountingDataBase","Company"), 				//Save the company name
		"address":banDoc.info("AccountingDataBase","Address1"), 			//Save the address																	//Save details
		"nation":banDoc.info("AccountingDataBase","Country"), 				//Save the country
		"telephone":banDoc.info("AccountingDataBase","Phone"), 				//Save the phone number
		"zip":banDoc.info("AccountingDataBase","Zip"), 						//Save the zip code
		"city":banDoc.info("AccountingDataBase","City"),					//Save the city
		"grColumn" : "Gr",													//Save the GR column (Gr1 or Gr2)
		"rounding" : 2,														//Speficy the rounding type		
		"formatNumber":true, 												//Choose if format number or not
		"groupByFormId":true, 												//Use form id to group accounts
		"groupByGr":false, 													//Use group accounts
	};
}

function loadForm() {

	/*
		PROFIT & LOSS
	*/
	GR = {};
	for (i = 1; i < 33; i++) {
		GR[i] = "损益表" + i;
		}
	form.push({"id":GR[1], "gr":"500", "bClass":"4", "description":"一、营业收入"});
	form.push({"id":GR[2], "gr":"530", "bClass":"3", "description":"减:营业成本"});
	form.push({"id":GR[3], "gr":"531", "bClass":"3", "description":"营业税金及附加"});
	form.push({"id":GR[4], "gr":"", "bClass":"3", "description":"其中:消费税"});
	form.push({"id":GR[5], "gr":"", "bClass":"3", "description":"营业税"});
	form.push({"id":GR[6], "gr":"", "bClass":"3", "description":"城市维护建设税"});
	form.push({"id":GR[7], "gr":"", "bClass":"3", "description":"资源税"});
	form.push({"id":GR[8], "gr":"", "bClass":"3", "description":"土地增值税"});
	form.push({"id":GR[9], "gr":"", "bClass":"3", "description":"城镇土地使用税、房产税、车船税、印花税"});
	form.push({"id":GR[10], "gr":"", "bClass":"3", "description":"教育费附加、矿产资源补偿费、排污费"});
	form.push({"id":GR[11], "gr":"540", "bClass":"3", "description":"销售费用"});
	form.push({"id":GR[12], "gr":"", "bClass":"3", "description":"其中:商品维修费"});
	form.push({"id":GR[13], "gr":"", "bClass":"3", "description":"广告费和业务宣传费"});
	form.push({"id":GR[14], "gr":"542", "bClass":"3", "description":"管理费用"});
	form.push({"id":GR[15], "gr":"", "bClass":"3", "description":"其中:开办费"});
	form.push({"id":GR[16], "gr":"540308", "bClass":"3", "description":"业务招待费"});
	form.push({"id":GR[17], "gr":"", "bClass":"3", "description":"研究费用"});
	form.push({"id":GR[18], "gr":"541", "bClass":"3", "description":"财务费用"});
	form.push({"id":GR[19], "gr":"", "bClass":"3", "description":"其中:利息费用(收入以&ndash;填列)"});
	form.push({"id":GR[20], "gr":"510", "bClass":"3", "description":"加：投资收益(亏损以-填列)"});
	sum = GR[1] + ";" + GR[2] + ";" + GR[3] + ";" + GR[4] + ";" + GR[5] + ";" + GR[6] + ";";  
	sum += GR[7] + ";" + GR[8] + ";" + GR[9] + ";" + GR[10] + ";" + GR[11] + ";" + GR[12] + ";";  
	sum += GR[19] + ";" + GR[20];  
	sum = GR[1] + ";" + GR[2] + ";" + GR[3] + ";" + GR[4] + ";" + GR[5] + ";" + GR[6] + ";";  
	sum = GR[1] + ";" + GR[2] + ";" + GR[3] + ";" + GR[4] + ";" + GR[5] + ";" + GR[6] + ";";  
	form.push({"id":GR[21], "description":"二、营业利润(亏损以-号填列)", "sum":sum });
	form.push({"id":GR[22], "gr":"520", "bClass":"3", "description":"加:营业外收入"});
	form.push({"id":GR[23], "gr":"", "bClass":"3", "description":"其中:政府补助"});
	form.push({"id":GR[24], "gr":"550", "bClass":"3", "description":"减:营业外支出"});
	form.push({"id":GR[25], "gr":"", "bClass":"3", "description":"其中:坏账损失"});
	form.push({"id":GR[26], "gr":"", "bClass":"3", "description":"无法收回的长期债券投资损失"});
	form.push({"id":GR[27], "gr":"", "bClass":"3", "description":"无法收回的长期股权投资损失"});
	form.push({"id":GR[28], "gr":"", "bClass":"3", "description":"自然灾害等不可抗力因素造成的损失"});
	form.push({"id":GR[29], "gr":"", "bClass":"3", "description":"税收滞纳金"});
	sum = GR[21] + ";" + GR[22] + ";-" + GR[24];  
	form.push({"id":GR[30], "description":"三、利润总额(亏损总额以-号填列)", "sum":sum});
	form.push({"id":GR[31], "gr":"560", "bClass":"3", "description":"减:所得税费用"});
	sum = GR[30] + ";-" + GR[31] ;  
	form.push({"id":GR[32], "description":"四、净利润(净亏损以-号填列)", "sum": sum});

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

	//Every time the script is executed we clear the messages in banana
	Banana.document.clearMessages();


	//Open a dialog window asking to select an item from the list
	//var itemSelected = Banana.Ui.getItem('Input', 'Choose a value', ["1. 资产负债表","2. 损益表"], 0, false);

	
	//Function call to manage and save user settings about the period date
	var dateform = getPeriodSettings();

	//Check if user has entered a period date.
	//If yes, we can get all the informations we need, process them and finally create the report.
	//If not, the script execution will be stopped immediately.
	if (dateform) {

		//Function call to create the report
		var report = createProfitReport(Banana.document, dateform.selectionStartDate, dateform.selectionEndDate);

		//Print the report
		var stylesheet = createStylesheet();
		Banana.Report.preview(report, stylesheet);	
	}
}


//The purpose of this function is to do some operations before the values are converted
function postProcessAmounts(banDoc) {}


//Function that create the report
function createProfitReport(banDoc, startDate, endDate) {

	/** 1. CREATE AND LOAD THE PARAMETERS AND THE FORM */
	loadParam(banDoc, startDate, endDate);
	loadForm();
	
	/** 2. EXTRACT THE DATA, CALCULATE AND LOAD THE BALANCES */
	loadBalancesFromBanana();
	
	/** 3. CALCULATE THE TOTALS */
	calcFormTotals(["amount"]);
	calcFormTotals(["currentMonthAmount"]);
	
	/** 4. DO SOME OPERATIONS BEFORE CONVERTING THE VALUES */
	//postProcessAmounts(banDoc);
	
	/** 5. CONVERT ALL THE VALUES */
	formatValues(["amount"]);
	formatValues(["currentMonthAmount"]);
	


	/** START PRINT... */
	/** ------------------------------------------------------------------------------------------------------------- */

	//Create a report.
	var report = Banana.Report.newReport(param.reportName);

	//Table with basic informations
	var table = report.addTable("table");

	report.addParagraph("利润表(适用执行小企业会计准则的企业) - (ALFA VERSION)", "bold"); //Balance sheet (application of accounting standards for small business enterprise)
	report.addParagraph(" ", "");

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


	//----------------------------------------------------------------------------------------------//
	// var table1 = report.addTable("table1");	
	// var tableHeader = table1.getHeader();
	tableRow = table.addRow();	
	tableRow.addCell("项   目", "valueTitle bold bordersTitle border-left border-top border-bottom center", 1); //Project
	tableRow.addCell("行次", "valueTitle bold bordersTitle border-top border-bottom center", 1); 				//Line no
	tableRow.addCell("本年累计金额", "valueTitle bold bordersTitle border-top border-bottom center", 1); 			//YTD amount
	tableRow.addCell("本月金额", "valueTitle bold bordersTitle border-top border-bottom center", 1); 				//this month amount


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[1], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[1], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[1], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[1], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[2], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[2], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[2], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[2], "currentMonthAmount"), "borders right padding-right", 1);

	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[3], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[3], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[3], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[3], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[4], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[4], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[4], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[4], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[5], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[5], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[5], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[5], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[6], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[6], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[6], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[6], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[7], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[7], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[7], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[7], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[8], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[8], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[8], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[8], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[9], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[9], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[9], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[9], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[10], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[10], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[10], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[10], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[11], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[11], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[11], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[11], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[12], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[12], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[12], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[12], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[13], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[13], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[13], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[13], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[14], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[14], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[14], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[14], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[15], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[15], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[15], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[15], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[16], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[16], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[16], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[16], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[17], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[17], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[17], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[17], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[18], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[18], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[18], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[18], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[19], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[19], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[19], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[19], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[20], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[20], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[20], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[20], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[21], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[21], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[21], "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, GR[21], "currentMonthAmount"), "borders right valueTotal padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[22], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[22], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[22], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[22], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[23], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[23], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[23], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[23], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[24], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[24], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[24], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[24], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[25], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[25], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[25], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[25], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[26], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[26], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[26], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[26], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[27], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[27], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[27], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[27], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[28], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[28], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[28], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[28], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[29], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[29], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[29], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[29], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[30], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[30], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[30], "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, GR[30], "currentMonthAmount"), "borders right valueTotal padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[31], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[31], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[31], "amount"), "borders right padding-right", 1);
	tableRow.addCell(getValue(form, GR[31], "currentMonthAmount"), "borders right padding-right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, GR[32], "description"), "borders", 1);
	tableRow.addCell(getValue(form, GR[32], "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, GR[32], "amount"), "borders right valueTotal padding-right", 1);
	tableRow.addCell(getValue(form, GR[32], "currentMonthAmount"), "borders right valueTotal padding-right", 1);

	/** END PRINT ... */
	/** ------------------------------------------------------------------------------------------------------------- */

	//Return the final report
	return report;
}



//Function that creates a string that contains all the accounts belonging to the same Gr1
function getAccountsByGr1(formId) {

	var strAccounts = [];
	var tAccounts = Banana.document.table("Accounts");
	var lenAcc = tAccounts.rowCount;
	var lenForm = form.length;

	for (var i = 0; i < lenAcc; i++) {
		var tRow = tAccounts.row(i);
		var tGr1 = tRow.value("Gr1");
		if (tGr1 == formId) {
			strAccounts.push(tRow.value("Account"));
		}
	}
	return strAccounts.join("|");;
}



function loadBalancesFromBanana() {

	for (var i in form) {
	
		if (form[i]["bClass"]) {

			var bClass = form[i]["bClass"];
			var currentBal;
			var currentBal1;

			if (param.groupByFormId) {
				//Sum the amounts of opening, debit, credit, total and balance for all the transactions belonging to the same Gr1	
				//strAccounts = "1000|1001|1111"
				var formId = form[i]["id"];
				var strAccounts = getAccountsByGr1(formId);
				currentBal = Banana.document.currentBalance(strAccounts, Banana.document.info("AccountingDataBase","OpeningDate"), param["endDate"]);
				currentBal1 = Banana.document.currentBalance(strAccounts, param["startDate"], param["endDate"]);
			}
			else if (param.groupByGr) {
				//Sum the amounts of opening, debit, credit, total and balance for all transactions for this accounts
				// gr = "Gr=1113|111"
				currentBal = Banana.document.currentBalance("Gr="+form[i]["gr"], Banana.document.info("AccountingDataBase","OpeningDate"), param["endDate"]);
				currentBal1 = Banana.document.currentBalance("Gr="+form[i]["gr"], param["startDate"], param["endDate"]);
			}
			
			form[i]["amount"] = "";
			form[i]["opening"] = "";
			
			//The "bClass" decides which value to use
			if (bClass === "0") {
				form[i]["amount"] = currentBal.balance;
			}
			else if (bClass === "1") {
				form[i]["amount"]  = currentBal.balance;
			}
			else if (bClass === "2") {
				form[i]["amount"]  = Banana.SDecimal.invert(currentBal.balance);
			}
			else if (bClass === "3") {
				form[i]["amount"]  = currentBal.total;
				form[i]["currentMonthAmount"]  = currentBal1.total;
			}
			else if (bClass === "4") {
				form[i]["amount"]  = Banana.SDecimal.invert(currentBal.total);
				form[i]["currentMonthAmount"]  = Banana.SDecimal.invert(currentBal1.total);
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
//When calling this function, it's necessary to specify the form (the structure), the object ID, and the field (parameter) needed.
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
    pageStyle.setAttribute("margin", "10mm 10mm 10mm 20mm");

    stylesheet.addStyle("body", "font-size: 10pt; font-family: Helvetica");
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
	//style.setAttribute("padding-bottom", "5px"); 
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
