function test() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var rangeList = activeSheet.getRangeList(['A1:B4', 'D1:E4']);
  rangeList.activate();

  var selection = activeSheet.getSelection();
  // Current Cell: D1
  Logger.log('Current Cell: ' + selection.getCurrentCell().getA1Notation());
  // Active Range: D1:E4
  Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());
  // Active Ranges: A1:B4, D1:E4
  var ranges =  selection.getActiveRangeList().getRanges();
  for (var i = 0; i < ranges.length; i++) {
    Logger.log('Active Ranges: ' + ranges[i].getA1Notation());
  }
  Logger.log('Active Sheet: ' + selection.getActiveSheet().getName());
}

function system_sheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = ss.getSheets()[4];
    for(var i = 2;i<=9;i++) {
      var sheetName = source.getRange(i,1).getValue();
      var yourNewSheet = ss.getSheetByName(sheetName);
      if (yourNewSheet != null) {
        ss.deleteSheet(yourNewSheet);
      } else{
        yourNewSheet = ss.insertSheet(sheetName,ss.getNumSheets());
        template(sheetName)

      }
    }
}

function getFirstEmptyRow() {
  var spr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Catalogo");
  var column = spr.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) {
    ct++;
  }
  return (ct);
}

function cataloga() {
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Sheet");
  var destination = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Catalogo");
  row = getFirstEmptyRow()
  row += 1
  source.getRange("D3").copyTo(destination.getRange(row,1),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D3").setValue(source.getRange("D3").getValue()+1)
  source.getRange("G3:H3").getMergedRanges()[0].copyTo(destination.getRange(row,2),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G3:H3").getMergedRanges()[0].setValue("")
  source.getRange("K3").copyTo(destination.getRange(row,3),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K3").setValue("")
  source.getRange("N3").copyTo(destination.getRange(row,4),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N3").setValue("")
  source.getRange("D5:D6").copyTo(destination.getRange(row,5),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D5:D6").setValue("")
  source.getRange("D7").copyTo(destination.getRange(row,6),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D7").setValue("")
  source.getRange("D8").copyTo(destination.getRange(row,7),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D8").setValue("")
  source.getRange("G5:H6").copyTo(destination.getRange(row,8),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G5:H6").setValue("")
  source.getRange("G7:H7").copyTo(destination.getRange(row,9),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G7:H7").setValue("")
  source.getRange("G8:H8").copyTo(destination.getRange(row,10),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G8:H8").setValue("")
  source.getRange("K5:K6").copyTo(destination.getRange(row,11),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K5:K6").setValue("")
  source.getRange("K7").copyTo(destination.getRange(row,12),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K7").setValue("")
  source.getRange("K8").copyTo(destination.getRange(row,13),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K8").setValue("")
  source.getRange("D10:K11").copyTo(destination.getRange(row,14),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D10:K11").setValue("")
  source.getRange("D14:F14").copyTo(destination.getRange(row,15),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D14:F14").setValue("")
  source.getRange("D17").copyTo(destination.getRange(row,16),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D17").setValue("")
  source.getRange("F17:G17").copyTo(destination.getRange(row,17),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("F17:G17").setValue("")
  source.getRange("I14:K17").copyTo(destination.getRange(row,18),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("I14:K17").setValue("")
  source.getRange("D19:K20").copyTo(destination.getRange(row,19),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D19:K20").setValue("")
  source.getRange("D23:F23").copyTo(destination.getRange(row,20),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D23:F23").setValue("")
  source.getRange("D25:F25").copyTo(destination.getRange(row,21),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D25:F25").setValue("")
  source.getRange("I23:K25").copyTo(destination.getRange(row,22),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("I23:K25").setValue("")
  source.getRange("D27").copyTo(destination.getRange(row,23),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D27").setValue("")
  source.getRange("D28").copyTo(destination.getRange(row,24),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D28").setValue("")
  source.getRange("D29").copyTo(destination.getRange(row,25),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D29").setValue("")
  source.getRange("G27:H27").copyTo(destination.getRange(row,26),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G27:H27").setValue("")
  source.getRange("G28:H28").copyTo(destination.getRange(row,27),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G28:H28").setValue("")
  source.getRange("G29:H29").copyTo(destination.getRange(row,28),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G29:H29").setValue("")
  source.getRange("K27").copyTo(destination.getRange(row,29),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K27").setValue("")
  source.getRange("K28").copyTo(destination.getRange(row,30),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K28").setValue("")
  source.getRange("K29").copyTo(destination.getRange(row,31),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K29").setValue("")
  source.getRange("N5:N6").copyTo(destination.getRange(row,32),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N5:N6").setValue("")
  source.getRange("N7").copyTo(destination.getRange(row,33),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N7").setValue("")
  source.getRange("N8").copyTo(destination.getRange(row,34),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N8").setValue("")
  source.getRange("N9").copyTo(destination.getRange(row,35),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N9").setValue("")
  source.getRange("N12:N13").copyTo(destination.getRange(row,36),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N12:N13").setValue("")
  source.getRange("N14").copyTo(destination.getRange(row,37),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N14").setValue("")
  source.getRange("N15:N16").copyTo(destination.getRange(row,38),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N15:N16").setValue("")
  source.getRange("N17").copyTo(destination.getRange(row,39),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N17").setValue("")
  source.getRange("N19").copyTo(destination.getRange(row,40),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N19").setValue("")
  source.getRange("N20").copyTo(destination.getRange(row,41),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N20").setValue("")
  source.getRange("N21:N22").copyTo(destination.getRange(row,42),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N21:N22").setValue("")
  source.getRange("N23").copyTo(destination.getRange(row,43),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N23").setValue("")
  source.getRange("N26").copyTo(destination.getRange(row,44),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N26").setValue("")
  source.getRange("N28").copyTo(destination.getRange(row,45),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N28").setValue("")
  source.getRange("N29").copyTo(destination.getRange(row,46),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N29").setValue("")
}

function aggiornaimport() {
  var destination = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interfaccia Bello");
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Catalogo");
  var number = destination.getRange("D3").getValue();
  for (var i = 2;i<=source.getMaxRows();i = i+1)  {
    check = source.getRange(i,1).getValue();
    if(check == number) {
      row = number
      break;
    }
  }
  source.getRange(row,1).copyTo(destination.getRange("D3"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,2).copyTo(destination.getRange("G3:H3").getMergedRanges()[0],SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,3).copyTo(destination.getRange("K3"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,4).copyTo(destination.getRange("N3"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,5).copyTo(destination.getRange("D5:D6"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,6).copyTo(destination.getRange("D7"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,7).copyTo(destination.getRange("D8"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,8).copyTo(destination.getRange("G5:H6"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,9).copyTo(destination.getRange("G7:H7"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,10).copyTo(destination.getRange("G8:H8"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,11).copyTo(destination.getRange("K5:K6"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,12).copyTo(destination.getRange("K7"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,13).copyTo(destination.getRange("K8"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,14).copyTo(destination.getRange("D10:K11"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,15).copyTo(destination.getRange("D14:F14"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,16).copyTo(destination.getRange("D17"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,17).copyTo(destination.getRange("F17:G17"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,18).copyTo(destination.getRange("I14:K17"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,19).copyTo(destination.getRange("D19:K20"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,20).copyTo(destination.getRange("D23:F23"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,21).copyTo(destination.getRange("D25:F25"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,22).copyTo(destination.getRange("I23:K25"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,23).copyTo(destination.getRange("D27"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,24).copyTo(destination.getRange("D28"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,25).copyTo(destination.getRange("D29"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,26).copyTo(destination.getRange("G27:H27"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,27).copyTo(destination.getRange("G28:H28"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,28).copyTo(destination.getRange("G29:H29"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,29).copyTo(destination.getRange("K27"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,30).copyTo(destination.getRange("K28"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,31).copyTo(destination.getRange("K29"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,32).copyTo(destination.getRange("N5:N6"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,33).copyTo(destination.getRange("N7"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,34).copyTo(destination.getRange("N8"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,35).copyTo(destination.getRange("N9"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,36).copyTo(destination.getRange("N12:N13"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,37).copyTo(destination.getRange("N14"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,38).copyTo(destination.getRange("N15:N16"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,39).copyTo(destination.getRange("N17"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,40).copyTo(destination.getRange("N19"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,41).copyTo(destination.getRange("N20"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,42).copyTo(destination.getRange("N21:N22"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,43).copyTo(destination.getRange("N23"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,44).copyTo(destination.getRange("N26"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,45).copyTo(destination.getRange("N28"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange(row,46).copyTo(destination.getRange("N29"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
}

function aggiornabottone()  {
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interfaccia Bello");
  var destination = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Catalogo");
  var number = source.getRange("D3").getValue();
  for (var i = 2;i<=destination.getMaxRows();i = i+1)  {
    check = destination.getRange(i,1).getValue();
    if(check == number) {
      row = number
      break;
    }
  }
  source.getRange("D3").copyTo(destination.getRange(row,1),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D3").setValue("")
  source.getRange("G3:H3").getMergedRanges()[0].copyTo(destination.getRange(row,2),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G3:H3").getMergedRanges()[0].setValue("")
  source.getRange("K3").copyTo(destination.getRange(row,3),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K3").setValue("")
  source.getRange("N3").copyTo(destination.getRange(row,4),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N3").setValue("")
  source.getRange("D5:D6").copyTo(destination.getRange(row,5),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D5:D6").setValue("")
  source.getRange("D7").copyTo(destination.getRange(row,6),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D7").setValue("")
  source.getRange("D8").copyTo(destination.getRange(row,7),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D8").setValue("")
  source.getRange("G5:H6").copyTo(destination.getRange(row,8),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G5:H6").setValue("")
  source.getRange("G7:H7").copyTo(destination.getRange(row,9),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G7:H7").setValue("")
  source.getRange("G8:H8").copyTo(destination.getRange(row,10),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G8:H8").setValue("")
  source.getRange("K5:K6").copyTo(destination.getRange(row,11),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K5:K6").setValue("")
  source.getRange("K7").copyTo(destination.getRange(row,12),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K7").setValue("")
  source.getRange("K8").copyTo(destination.getRange(row,13),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K8").setValue("")
  source.getRange("D10:K11").copyTo(destination.getRange(row,14),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D10:K11").setValue("")
  source.getRange("D14:F14").copyTo(destination.getRange(row,15),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D14:F14").setValue("")
  source.getRange("D17").copyTo(destination.getRange(row,16),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D17").setValue("")
  source.getRange("F17:G17").copyTo(destination.getRange(row,17),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("F17:G17").setValue("")
  source.getRange("I14:K17").copyTo(destination.getRange(row,18),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("I14:K17").setValue("")
  source.getRange("D19:K20").copyTo(destination.getRange(row,19),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D19:K20").setValue("")
  source.getRange("D23:F23").copyTo(destination.getRange(row,20),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D23:F23").setValue("")
  source.getRange("D25:F25").copyTo(destination.getRange(row,21),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D25:F25").setValue("")
  source.getRange("I23:K25").copyTo(destination.getRange(row,22),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("I23:K25").setValue("")
  source.getRange("D27").copyTo(destination.getRange(row,23),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D27").setValue("")
  source.getRange("D28").copyTo(destination.getRange(row,24),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D28").setValue("")
  source.getRange("D29").copyTo(destination.getRange(row,25),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("D29").setValue("")
  source.getRange("G27:H27").copyTo(destination.getRange(row,26),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G27:H27").setValue("")
  source.getRange("G28:H28").copyTo(destination.getRange(row,27),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G28:H28").setValue("")
  source.getRange("G29:H29").copyTo(destination.getRange(row,28),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("G29:H29").setValue("")
  source.getRange("K27").copyTo(destination.getRange(row,29),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K27").setValue("")
  source.getRange("K28").copyTo(destination.getRange(row,30),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K28").setValue("")
  source.getRange("K29").copyTo(destination.getRange(row,31),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("K29").setValue("")
  source.getRange("N5:N6").copyTo(destination.getRange(row,32),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N5:N6").setValue("")
  source.getRange("N7").copyTo(destination.getRange(row,33),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N7").setValue("")
  source.getRange("N8").copyTo(destination.getRange(row,34),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N8").setValue("")
  source.getRange("N9").copyTo(destination.getRange(row,35),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N9").setValue("")
  source.getRange("N12:N13").copyTo(destination.getRange(row,36),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N12:N13").setValue("")
  source.getRange("N14").copyTo(destination.getRange(row,37),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N14").setValue("")
  source.getRange("N15:N16").copyTo(destination.getRange(row,38),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N15:N16").setValue("")
  source.getRange("N17").copyTo(destination.getRange(row,39),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N17").setValue("")
  source.getRange("N19").copyTo(destination.getRange(row,40),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N19").setValue("")
  source.getRange("N20").copyTo(destination.getRange(row,41),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N20").setValue("")
  source.getRange("N21:N22").copyTo(destination.getRange(row,42),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N21:N22").setValue("")
  source.getRange("N23").copyTo(destination.getRange(row,43),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N23").setValue("")
  source.getRange("N26").copyTo(destination.getRange(row,44),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N26").setValue("")
  source.getRange("N28").copyTo(destination.getRange(row,45),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N28").setValue("")
  source.getRange("N29").copyTo(destination.getRange(row,46),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
  source.getRange("N29").setValue("")
}

function event(e) {
  var event = e.range.getA1Notation();
  if(event == "D3") {
    aggiornaimport();
  }
}