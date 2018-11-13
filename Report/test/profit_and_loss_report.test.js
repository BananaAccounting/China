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
// @id = ch.banana.china.reportprofitloss.test
// @api = 1.0
// @pubdate = 2018-11-06
// @publisher = Banana.ch SA
// @description.en = Profit and Loss
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @outputformat = none
// @inputdataform = none
// @timeout = -1
// @includejs = ../profit_and_loss_report.js

// Register test case to be executed
Test.registerTestCase(new ReportProfitLoss());

// Here we define the class, the name of the class is not important
function ReportProfitLoss() {

}

// This method will be called at the beginning of the test case
ReportProfitLoss.prototype.initTestCase = function() {

}

// This method will be called at the end of the test case
ReportProfitLoss.prototype.cleanupTestCase = function() {

}

// This method will be called before every test method is executed
ReportProfitLoss.prototype.init = function() {

}

// This method will be called after every test method is executed
ReportProfitLoss.prototype.cleanup = function() {

}

// Generate the expected (correct) file
ReportProfitLoss.prototype.testBananaApp = function() {

  //Open the banana document
  var banDoc = Banana.application.openDocument("file:script/../test/testcases/template2.ac2");
  Test.assert(banDoc);
  
  this.report_test(banDoc, "2018-01-01", "2018-12-31", "Whole year report");
  // this.report_test(banDoc, "2018-01-01", "2018-06-30", "First semester report");
  // this.report_test(banDoc, "2018-07-01", "2018-12-31", "Second semester report");
  // this.report_test(banDoc, "2018-01-01", "2018-03-31", "First quarter report");
  // this.report_test(banDoc, "2018-04-01", "2018-06-30", "Second quarter report");
  // this.report_test(banDoc, "2018-07-01", "2018-09-30", "Third quarter report");
  // this.report_test(banDoc, "2018-10-01", "2018-12-31", "Fourth quarter report");

}

ReportProfitLoss.prototype.report_test = function(banDoc, startDate, endDate, reportName) {
  var report = createProfitReport(banDoc, startDate, endDate);
  Test.logger.addReport(reportName, report);
}

