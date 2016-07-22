var nNumberPattern = /^[nN]\d{8}$/
var nNumberWrongPattern = /^[nN][*:\-.\/]\d{8}$/
var nNumberShortLongPattern = /^[nN]\d{7,9}$/
var thisSS = SpreadsheetApp.getActiveSpreadsheet()