var date = 'A2:A';
var pay = 'F2:F';
var monthlyWageDst = 'L10';

function countMonthlyWages(e){
  if (e.source.getActiveRange().getColumn() == e.source.getRange(pay).getColumn()) {
    var name = e.source.getActiveSheet().getName();
    switch (name){
      case "tolik2018":
      case "tolik2019":
      case "tolik2020":
      case "tolik2021":
      case "tolik2022":
      case "tolik2023":
      case "tolik2024":
        monthlyWages(name);
        break;
    }
  }
}

function monthlyWages(sheet){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);
  var dates = activeSpreadsheet.getRange(date);
  var latestDate = 1;
  var wages = activeSpreadsheet.getRange(pay);
  var latestPay = 1;
  // trimming the wages range
  for (var i = 1; i < wages.getLastRow(); i ++){
    if (!wages.getCell(i, 1).isBlank()) latestPay = i;
  }
  // trimming the dates range
  for (var i = latestPay; i > 0; i --){
    if (!dates.getCell(i, 1).isBlank()){
      latestDate = i;
      break;
    }
  }
  // determinig the date 30 days before the latest date
  var endMonth = new Date(dates.getCell(latestDate, 1).getValue());
  var startMonth;
  var startDate;
  for (var i = latestDate - 1; i > 0; i --){
    if (!dates.getCell(i, 1).isBlank()){
      startMonth = new Date (dates.getCell(i, 1).getValue());
      startDate = i;
      if (endMonth.getTime() - startMonth.getTime() >= 30 * 24 * 60 * 60 * 1000){
      //  startDate = i;
        break;
      }
    }
  }
  // summing up the wages starting from the line of the date 30 days before the latest date up to the latest wages
  var retWages = 0;
  for (var i = startDate; i <= latestPay; i ++) retWages = +retWages + +wages.getCell(i, 1).getValue();
  activeSpreadsheet.getRange(monthlyWageDst).setValue(retWages);
}
