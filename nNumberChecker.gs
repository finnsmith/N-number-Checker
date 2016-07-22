function onInstall() {
  onOpen();
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem("Show N-number Checker", "showSidebar")
  .addToUi();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile("nNumberCheckerSidebar")
  .evaluate()
  .setTitle("N-number Checker")
  SpreadsheetApp.getUi().showSidebar(html)
}

function checkNumbers() {
  var sheets = thisSS.getSheets()
  
  for (var S in sheets) {
    var thisSheet = sheets[S]
    var sheetName = thisSheet.getName()
    var maxWidth = 0
    var origData = thisSheet.getDataRange().getValues()
    
    for (var R in origData) {
      var row = origData[R]
      var rowNumber = parseInt(R) + 1
      checkRow(sheetName, row, rowNumber)
    }
  } 
}

function checkRow(sheetName, row, rowNumber) {
  for (C in row) {
    var cell = row[C]
    
    var noWhiteSpace = cell.toString().replace(/\s+/g, '')
    switch (true) {
        
      case nNumberPattern.test(noWhiteSpace):
        cell  = "N" + noWhiteSpace.slice(1)
        thisSS.getSheetByName(sheetName).getRange(rowNumber, (parseInt(C) + 1)).setValue(cell)
        break;
        
      case nNumberShortLongPattern.test(noWhiteSpace):
        thisSS.getSheetByName(sheetName).getRange(rowNumber, (parseInt(C) + 1)).setBackground('#58068c');
        thisSS.getSheetByName(sheetName).getRange(rowNumber, (parseInt(C) + 1)).setFontColor('#ffffff')
        break;
        
      case nNumberWrongPattern.test(noWhiteSpace):
        cell = "N" + noWhiteSpace.slice(2)
        thisSS.getSheetByName(sheetName).getRange(rowNumber, (parseInt(C) + 1)).setValue(cell)
        break;
        
      case nNumberPattern.test("N" + noWhiteSpace):
        cell = "N" + noWhiteSpace
        thisSS.getSheetByName(sheetName).getRange(rowNumber, (parseInt(C) + 1)).setValue(cell)
        break;
           
    }
  }
}