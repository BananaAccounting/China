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


// @id = ch.banana.apps.import.csvstatement.pinganbank.test
// @api = 1.0
// @pubdate = 2017-11-24
// @publisher = Banana.ch SA
// @description = <TEST PingAnBank.csv 平安银行>
// @task = app.command
// @doctype = *.*
// @docproperties =
// @outputformat = none
// @inputdataform = none
// @includejs = ../import.csvstatement.pinganbank.js
// @timeout = -1


/*

SUMMARY
-------
Javascript test
1. Open the .ac2 file
2. Execute the .js script
3. Save the report



virtual void addTestBegin(const QString& key, const QString& comment = QString());
virtual void addTestEnd();

virtual void addSection(const QString& key);
virtual void addSubSection(const QString& key);
virtual void addSubSubSection(const QString& key);

virtual void addComment(const QString& comment);
virtual void addInfo(const QString& key, const QString& value1, const QString& value2 = QString(), const QString& value3 = QString());
virtual void addFatalError(const QString& error);
virtual void addKeyValue(const QString& key, const QString& value, const QString& comment = QString());
virtual void addReport(const QString& key, QJSValue report, const QString& comment = QString());
virtual void addTable(const QString& key, QJSValue table, QStringList colXmlNames = QStringList(), const QString& comment = QString());

 **/

// Register test case to be executed
Test.registerTestCase(new ImportExcelCsvPingabankTest());

// Here we define the class, the name of the class is not important
function ImportExcelCsvPingabankTest() {}

// This method will be called at the beginning of the test case
ImportExcelCsvPingabankTest.prototype.initTestCase = function () {}

// This method will be called at the end of the test case
ImportExcelCsvPingabankTest.prototype.cleanupTestCase = function () {}

// This method will be called before every test method is executed
ImportExcelCsvPingabankTest.prototype.init = function () {}

// This method will be called after every test method is executed
ImportExcelCsvPingabankTest.prototype.cleanup = function () {}

ImportExcelCsvPingabankTest.prototype.testImport = function () {

	Test.logger.addKeyValue("key1", "test2");
	Test.logger.addComment("Test vImportExcelCsvPingabankTest");

	var file = Banana.IO.getLocalFile("file:script/testcases/PingAnBank.csv平安银行.csv")
		file.codecName = "GBK"; // Default is UTF-8
	var fileContent = file.read();
	if (file.errorString) {
		Test.logger.addComment(file.errorString);
		Test.logger.addComment("Test impossible to import");
		return;
	}
	//Banana.console.debug(fileContent);
	var tabSeparatedText = exec(fileContent);
	var report = createReportFromTabSeparatedText(tabSeparatedText);
	Test.logger.addReport("Report converted text", report);
}

function createReportFromTabSeparatedText(tabSeparatedText) {

	var report = Banana.Report.newReport("Report from tabSeparatedText");
	var table = report.addTable();
	var arrData = Banana.Converter.csvToArray(tabSeparatedText, '\t');
	if (!arrData) {
		Test.logger.addComment("Error no data in tabSeparatedText");
		return;
	}
	for (var i = 0; i < arrData.length; i++) {
		var tableRow = table.addRow();
		var values = arrData[i];
		if (values) {
			for (var j = 0; j < values.length; ++j) {
				tableRow.addCell(values[j]);
			}
		}
	}
	return report;
}

