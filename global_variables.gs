var as = SpreadsheetApp.getActiveSpreadsheet()
var respSheet = as.getSheetByName('回答')
var totalSheet = as.getSheetByName('概要')
var tempSheet = as.getSheetByName('テンプレート')

var respLR = respSheet.getLastRow()
var respLC = respSheet.getLastColumn()

var newInput = respSheet.getRange(respLR,1,1,respLC).getValues()
var inputDate, payer, paidDate, purpose, paidAmount, chargeYui, chargeHaru   
inputDate = newInput[0][0]
payer = newInput[0][1]
paidDate = newInput[0][2]
purpose = newInput[0][3]
paidAmount = newInput[0][4]
chargeYui = newInput[0][5]
chargeHaru = newInput[0][6]