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
// @id = ch.banana.addon.voucherchina.test
// @api = 1.0
// @pubdate = 2019-03-05
// @publisher = Banana.ch SA
// @description.en = Balance Sheet
// @task = app.command
// @doctype = 100.*;110.*;130.*
// @outputformat = none
// @inputdataform = none
// @timeout = -1
// @includejs = ../voucher_china.js

// Register test case to be executed
Test.registerTestCase(new ReportVoucher());

// Here we define the class, the name of the class is not important
function ReportVoucher() {

}

// This method will be called at the beginning of the test case
ReportVoucher.prototype.initTestCase = function() {

}

// This method will be called at the end of the test case
ReportVoucher.prototype.cleanupTestCase = function() {

}

// This method will be called before every test method is executed
ReportVoucher.prototype.init = function() {

}

// This method will be called after every test method is executed
ReportVoucher.prototype.cleanup = function() {

}

// Generate the expected (correct) file
ReportVoucher.prototype.testBananaApp = function() {

  //Open the banana document
  var banDoc = Banana.application.openDocument("file:script/../test/testcases/template1.ac2");
  Test.assert(banDoc);

  loadParam();

  /* Test 1 */
  var userParam = {};
  userParam.voucher = '01-12';
  userParam.approved = 'AAA';
  userParam.checked = 'BBB';
  userParam.entered = 'CCC';
  userParam.cashier = 'DDD';
  userParam.prepared = 'EEE';
  userParam.receiver = 'FFF';
  userParam.printenglish = true;
  this.report_test(banDoc, userParam, "Single Voucher report, chinese and english");

  /* Test 2 */
  var userParam = {};
  userParam.voucher = '01-01';
  userParam.approved = 'A';
  userParam.checked = 'B';
  userParam.entered = 'C';
  userParam.cashier = 'D';
  userParam.prepared = 'E';
  userParam.receiver = 'F';
  userParam.printenglish = false;
  this.report_test(banDoc, userParam, "Single Voucher report, chinese only");

  /* Test 3 */
  var userParam = {};
  userParam.voucher = '*';
  userParam.approved = 'Aaaa';
  userParam.checked = 'Bbbb';
  userParam.entered = 'Cccc';
  userParam.cashier = 'Dddd';
  userParam.prepared = 'Eeee';
  userParam.receiver = 'Ffff';
  userParam.printenglish = true;
  this.report_test(banDoc, userParam, "All Vouchers report, chinese and english");

  /* Test 4 */
  var userParam = {};
  userParam.voucher = '*';
  userParam.approved = 'Aa';
  userParam.checked = 'Bb';
  userParam.entered = 'Cc';
  userParam.cashier = 'Dd';
  userParam.prepared = 'Ee';
  userParam.receiver = 'Ff';
  userParam.printenglish = false;
  this.report_test(banDoc, userParam, "All Vouchers report, chinese only");
}

ReportVoucher.prototype.report_test = function(banDoc, userParam, reportName) {
  var report = printVoucher(banDoc, userParam);
  Test.logger.addReport(reportName, report);
}

