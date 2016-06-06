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
// @id = ch.banana.addon.profitlossreportchina2016
// @api = 1.0
// @pubdate = 2016-04-04
// @publisher = Banana.ch SA
// @description = 损益表 2016
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
	var date = new Date();
	param = {
		"reportName":"China - Profit/Loss Report 2016",						//Save the report's name
		//"bananaVersion":"Banana Accounting 8", 							//Save the version of Banana Accounting used
		//"scriptVersion":"script v. 2016-04-04",				 			//Save the version of the script
		//"fiscalNumber":banDoc.info("AccountingDataBase","FiscalNumber"),	//Save the fiscal number
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
		"formatNumber":true 												//Choose if format number or not
	};
}

function loadForm() {

	/*
		PROFIT & LOSS
	*/
	form.push({"id":"P1", "gr":"4", "bClass":"4", "description":"一、营业收入"});
	form.push({"id":"P2", "gr":"5", "bClass":"3", "description":"减:营业成本"});
	form.push({"id":"P3", "gr":"217", "bClass":"2", "description":"营业税金及附加"});
	form.push({"id":"P4", "gr":"", "bClass":"2", "description":"其中:消费税"});
	form.push({"id":"P5", "gr":"", "bClass":"2", "description":"营业税"});
	form.push({"id":"P6", "gr":"", "bClass":"2", "description":"城市维护建设税"});
	form.push({"id":"P7", "gr":"", "bClass":"2", "description":"资源税"});
	form.push({"id":"P8", "gr":"234", "bClass":"2", "description":"土地增值税"});
	form.push({"id":"P9", "gr":"236", "bClass":"2", "description":"城镇土地使用税、房产税、车船税、印花税"});
	form.push({"id":"P10", "gr":"515~518", "bClass":"3", "description":"教育费附加、矿产资源补偿费、排污费"});
	form.push({"id":"P11", "gr":"", "bClass":"3", "description":"销售费用"});
	form.push({"id":"P12", "gr":"", "bClass":"3", "description":"其中:商品维修费"});
	form.push({"id":"P13", "gr":"615~618", "bClass":"3", "description":"广告费和业务宣传费"});
	form.push({"id":"P14", "gr":"62", "bClass":"3", "description":"管理费用"});
	form.push({"id":"P15", "gr":"625~628", "bClass":"3", "description":"其中:开办费"});
	form.push({"id":"P16", "gr":"", "bClass":"3", "description":"业务招待费"});
	form.push({"id":"P17", "gr":"63", "bClass":"3", "description":"研究费用"});
	form.push({"id":"P18", "gr":"", "bClass":"3", "description":"财务费用"});
	form.push({"id":"P19", "gr":"751", "bClass":"3", "description":"其中:利息费用(收入以&ndash;填列)"});
	form.push({"id":"P20", "gr":"712", "bClass":"3", "description":"加：投资收益(亏损以-填列)"});
	form.push({"id":"P21", "description":"二、营业利润(亏损以-号填列)", "sum":"P1;P2;P3;P4;P5;P6;P7;P8;P9;P10;P11;P12;P13;P14;P15;P16;P17;P18;P19;P20"});
	form.push({"id":"P22", "gr":"71~74", "bClass":"3", "description":"加:营业外收入"});
	form.push({"id":"P23", "gr":"748", "bClass":"3", "description":"其中:政府补助"});
	form.push({"id":"P24", "gr":"75~78", "bClass":"3", "description":"减:营业外支出"});
	form.push({"id":"P25", "gr":"615~618", "bClass":"3", "description":"其中:坏账损失"});
	form.push({"id":"P26", "gr":"752", "bClass":"3", "description":"无法收回的长期债券投资损失"});
	form.push({"id":"P27", "gr":"", "bClass":"3", "description":"无法收回的长期股权投资损失"});
	form.push({"id":"P28", "gr":"788", "bClass":"3", "description":"自然灾害等不可抗力因素造成的损失"});
	form.push({"id":"P29", "gr":"", "bClass":"3", "description":"税收滞纳金"});
	form.push({"id":"P30", "description":"三、利润总额(亏损总额以-号填列)", "sum":"P22;P23;P24;P25;P26;P27;P28;P29"});
	form.push({"id":"P31", "gr":"811", "bClass":"3", "description":"减:所得税费用"});
	form.push({"id":"P32", "description":"四、净利润(净亏损以-号填列)", "sum":"P21;P30"});

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
	loadBalances();
	
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
	tableRow.addCell(getValue(form, "P1", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P1", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P1", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P1", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P2", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P2", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P2", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P2", "currentMonthAmount"), "borders right", 1);

	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P3", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P3", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P3", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P3", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P4", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P4", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P4", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P4", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P5", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P5", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P5", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P5", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P6", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P6", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P6", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P6", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P7", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P7", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P7", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P7", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P8", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P8", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P8", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P8", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P9", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P9", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P9", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P9", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P10", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P10", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P10", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P10", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P11", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P11", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P11", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P11", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P12", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P12", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P12", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P12", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P13", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P13", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P13", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P13", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P14", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P14", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P14", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P14", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P15", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P15", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P15", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P15", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P16", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P16", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P16", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P16", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P17", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P17", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P17", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P17", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P18", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P18", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P18", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P18", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P19", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P19", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P19", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P19", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P20", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P20", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P20", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P20", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P21", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P21", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P21", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "P21", "currentMonthAmount"), "borders right valueTotal", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P22", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P22", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P22", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P22", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P23", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P23", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P23", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P23", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P24", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P24", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P24", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P24", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P25", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P25", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P25", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P25", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P26", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P26", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P26", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P26", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P27", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P27", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P27", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P27", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P28", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P28", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P28", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P28", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P29", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P29", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P29", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P29", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P30", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P30", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P30", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "P30", "currentMonthAmount"), "borders right valueTotal", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P31", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P31", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P31", "amount"), "borders right", 1);
	tableRow.addCell(getValue(form, "P31", "currentMonthAmount"), "borders right", 1);


	tableRow = table.addRow();
	tableRow.addCell(getValue(form, "P32", "description"), "borders", 1);
	tableRow.addCell(getValue(form, "P32", "id").replace("P", ""), "borders", 1);
	tableRow.addCell(getValue(form, "P32", "amount"), "borders right valueTotal", 1);
	tableRow.addCell(getValue(form, "P32", "currentMonthAmount"), "borders right valueTotal", 1);

	/** END PRINT ... */
	/** ------------------------------------------------------------------------------------------------------------- */

	//Return the final report
	return report;
}







function loadBalances() {

	for (var i in form) {
	
		if (form[i]["bClass"]) {

			if (form[i]["gr"]) {

				var bClass = form[i]["bClass"];

				//Sum the amounts of opening, debit, credit, total and balance for all transactions for this accounts
				// gr = "Gr=1113|111"
				var currentBal = Banana.document.currentBalance("Gr="+form[i]["gr"], Banana.document.info("AccountingDataBase","OpeningDate"), param["endDate"]);
				var currentBal1 = Banana.document.currentBalance("Gr="+form[i]["gr"], param["startDate"], param["endDate"]);
				
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


// //This function adds a Footer to the report
// function addFooter(report, param) {
//    report.getFooter().addClass("footer");
//    var versionLine = report.getFooter().addText(param.bananaVersion + ", " + param.scriptVersion + ", ", "description");
//    //versionLine.excludeFromTest();
//    report.getFooter().addText(param.pageCounterText + " ", "description");
//    report.getFooter().addFieldPageNr();
// }


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
	style.setAttribute("padding-bottom", "2px");
	style.setAttribute("padding-top", "5px"); 


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






// function daysInMonth(month,year) {
//     return new Date(year, month, 0).getDate();
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