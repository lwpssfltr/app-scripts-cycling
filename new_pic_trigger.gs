var lookup_table = 'pics';
var name = 'C2:C';

function new_pic_trigger(e) {
    if (e.source.getActiveSheet().getName() == lookup_table) newPick(e);
}

function newPick(event) {
  var lookup_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(lookup_table);
  var names = lookup_sheet.getRange('A1:A');
  var urls = lookup_sheet.getRange('B1:B');
  var r = lookup_sheet.getActiveCell().getRow();
  if (r == 1) return;
  if (!lookup_sheet.getRange(r, 1).isBlank() && !lookup_sheet.getRange(r, 2).isBlank()) {
    var cmd = lookup_sheet.getRange(r, 3).getValue();
    switch (cmd){
      case "n":
        break;
      case "r":
        if (event.range.getColumn() == 1) {
          var name_old = "" + event.oldValue;
          var name_new = "" + event.value;
          var url_str = lookup_sheet.getRange(event.range.getRow(), 2).getValue();
          updatePics("tolik2022", name_old, name_new, url_str);
          updatePics("tolik2023", name_old, name_new, url_str);
          updatePics("tolik2024", name_old, name_new, url_str);
          //updatePics("tolik2020", name_old, name_new, url_str);
        }
        break;
      default:
        var name_str = lookup_sheet.getRange(event.range.getRow(), 1).getValue();
        var url_str = lookup_sheet.getRange(event.range.getRow(), 2).getValue();
        updatePics("tolik2022", name_str, name_str, url_str);
        updatePics("tolik2023", name_str, name_str, url_str);
        updatePics("tolik2024", name_str, name_str, url_str);
        // updatePics("tolik2018", name_str, name_str, url_str);
        //updatePics("tolik2020", name_str, name_str, url_str);
    }    
  }
}

function updatePics(sheet_name, name_value_1, name_value_2, url_value) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var name_r = sh.getRange(name);
  var regexp_name_old = "";
  for (var i = 0; i < name_value_1.length; i ++) {
    switch(name_value_1.charAt(i)) {
      case "[":
      case "]":
      case "(":
      case ")":
        regexp_name_old = regexp_name_old + "\\";
      default:
        regexp_name_old = regexp_name_old + name_value_1.charAt(i);
    }
  }
  var regexp_name_new = "";
  for (var i = 0; i < name_value_1.length; i ++) {
    switch(name_value_2.charAt(i)) {
      case "[":
      case "]":
      case "(":
      case ")":
        regexp_name_new = regexp_name_new + "\\";
      default:
        regexp_name_new = regexp_name_new + name_value_2.charAt(i);
    }
  }
  var old_name_regexp = new RegExp('.*" *' + regexp_name_old + ' *"|^ *' + regexp_name_old + ' *$', 'i');
  var new_name_regexp = new RegExp('.*" *' + regexp_name_new + ' *"|^ *' + regexp_name_new + ' *$', 'i');
  // Browser.msgBox(name_regexp + ", " + activeName.length);
  for (var i = 1; i < name_r.getLastRow(); i ++) {
    var cell = name_r.getCell(i, 1);
    if (!cell.isBlank()) {
      if (old_name_regexp.test(cell.getValue()) || new_name_regexp.test(cell.getValue())) {
        cell.setValue('=HYPERLINK("' + url_value + '"; "' + name_value_2 + '")');
      }
    }
  }
}
