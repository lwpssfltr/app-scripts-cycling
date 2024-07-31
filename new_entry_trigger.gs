var name = 'C2:C';
var lookup_table = 'pics';



function new_entry_trigger(e) {
  if (e.source.getActiveRange().getColumn() == e.source.getRange(name).getColumn()) {
    var sheet = e.source.getActiveSheet().getName();
    switch (sheet){
      case "tolik2018":
      case "tolik2019":
      case "tolik2020":
      case "tolik2021":
      case "tolik2022":
      case "tolik2023":
      case "tolik2024":
        newEntry(sheet);
        //Browser.msgBox('=HYPERLINK("http://sas.com"; "tunis (stx 7) [3]")');
        break;
    }
  }
}
function newEntry(sheet_name){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var names = activeSpreadsheet.getRange(name);
  var active_cell = activeSpreadsheet.getActiveCell();
  var active_name_raw = "" + active_cell.getValue();
  var active_name = active_name_raw.replace(/=hyperlink\(".*"; "/i, "").replace(/"\)/i, "");
  var lookup_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(lookup_table);
  var lookup_names_range = lookup_sheet.getRange("A2:A");
  var lookup_urls_range = lookup_sheet.getRange("B2:B");
  var commands_range = lookup_sheet.getRange("C2:C");
  var active_name_regexp = "";
  for (var i = 0; i < active_name.length; i ++) {
    switch(active_name.charAt(i)) {
      case "[":
      case "]":
      case "(":
      case ")":
        active_name_regexp = active_name_regexp + "\\";
      default:
        active_name_regexp = active_name_regexp + active_name.charAt(i);
    }
  }
  var name_regexp = new RegExp('^' + active_name_regexp + '$', 'i');
  for (var i = 1; i < lookup_names_range.getLastRow(); i ++) {
    var cell = lookup_names_range.getCell(i, 1);
    if (!cell.isBlank()) {
      if (name_regexp.test(cell.getValue())) {
        if (commands_range.getCell(i, 1).getValue() != 'n' && commands_range.getCell(i, 1) != 'r') {
          active_cell.setValue('=HYPERLINK("' + lookup_urls_range.getCell(i, 1).getValue() + '"; "' + lookup_names_range.getCell(i, 1).getValue() + '")');
          break;
        }
      }
    }
  }
}
