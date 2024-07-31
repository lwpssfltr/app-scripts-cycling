var date = 'A2:A';
var pay = 'F2:F';
var maxWageDst = 'L12';

function countMaxDailyWages(e) {
  if (e.source.getActiveRange().getColumn() == e.source.getRange(pay).getColumn()) {
    var name = e.source.getActiveSheet().getName();
    switch (name){
      case "tolik2019":
      case "tolik2018":
      case "tolik2020":
      case "tolik2021":
      case "tolik2022":
      case "tolik2023":
      case "tolik2024":
        doTheMath(name);
        break;
    }
  }
}

function doTheMath(sheet) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);
  var dates = activeSpreadsheet.getRange(date);
  var wages = activeSpreadsheet.getRange(pay);
  var latestPay = 1;
  // trimming the wages range
  for (var i = 1; i < wages.getLastRow(); i ++){
    if (!wages.getCell(i, 1).isBlank()){
      latestPay = i;
    }
  }
  // trimming the dates range
  var firstDate = 1;
  for (var i = 1; i < latestPay; i ++){
    if (!dates.getCell(i, 1).isBlank()) {
      firstDate = i;
      break;
    }
  }
  var maxDailyPay = 0;
  
  var flagNextDate = false;
  
  var dailyTemp = 0;
  for (var i = firstDate; i <= latestPay; i ++) {
    if (!dates.getCell(i + 1, 1).isBlank()) {
      flagNextDate = true;
    }
    dailyTemp += +wages.getCell(i, 1).getValue();
    if (flagNextDate || i == latestPay){
      if (dailyTemp > maxDailyPay) {
        maxDailyPay = dailyTemp;
      }
      dailyTemp = 0;
      flagNextDate = false;
    }
//    Logger.log('line: ' + i);
//    Logger.log('max daily pay: ' + maxDailyPay);
//    Logger.log('daily temp: ' + dailyTemp);
  }
  activeSpreadsheet.getRange(maxWageDst).setValue(maxDailyPay);
}
