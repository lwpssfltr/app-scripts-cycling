var nameRange = 'C2:C';
var statusRange = 'G2:G';
var highlightedRange = 'K18:L';
var uncheckedRange = 'N18:N'; 
var graduateRange = 'L14';
  
function rebuildBikeStats(e){
 if (e.source.getActiveRange().getColumn() == e.source.getRange(statusRange).getColumn()) {
   var name = e.source.getActiveSheet().getName();
    switch (name){
      case "tolik2019":
      case "tolik2018":
      case "tolik2020":
      case "tolik2021":
      case "tolik2022":
      case "tolik2023":
      case "tolik2024":
        buildLists(name);
        break;
    }
  }
}
function buildLists(sheet) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);
  var names = activeSpreadsheet.getRange(nameRange);
  var stati = activeSpreadsheet.getRange(statusRange);
  var uncheckedList = activeSpreadsheet.getRange(uncheckedRange);
  var highlightedList = activeSpreadsheet.getRange(highlightedRange);
  var graduateList = activeSpreadsheet.getRange(graduateRange);
  uncheckedList.setValue('');
  highlightedList.setValue('');
  graduateList.setValue('');
  var uncheckedListPointer = 1;
  var highlightedListPointer = 1;
  var graduateListCounter = 0;
  for (var i = 1; i <= stati.getHeight(); i ++) {
  //  Logger.log(names.getCell(i, 1).getValue() + ' ' + stati.getCell(i, 1).getBackground());
    switch (stati.getCell(i, 1).getBackground()) {
      case '#ff00ffff':
      case '#00ffff':
        highlightedList.getCell(highlightedListPointer, 1).setValue(names.getCell(i, 1).getValue());
        highlightedList.getCell(highlightedListPointer, 2).setValue(stati.getCell(i, 1).getValue());
        highlightedListPointer ++;
        break;
      case '#ffff00':
      case '#ffffff00':
        uncheckedList.getCell(uncheckedListPointer, 1).setValue(names.getCell(i, 1).getValue());
        uncheckedListPointer ++;
        break;
      case '#00ff00':
      case '#ff00ff00':
        graduateListCounter ++;
    }
  }
  graduateList.getCell(1, 1).setValue(graduateListCounter);
}
