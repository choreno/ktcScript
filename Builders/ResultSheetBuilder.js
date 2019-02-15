
function ResultSheetBuilder() {

  resultSheet = new ResultSheet();

  resultSheet.initialize();
  resultSheet.addResultData();
  resultSheet.addSparkLineData();

  resultSheet.deleteRemainedRowsAndColumns();
  resultSheet.setCellCenterAndMiddle();

  resultSheet.setColumnWidth(55); //column width
  resultSheet.mergeCells(); //merge date cells

  resultSheet.setNamedRange();

}

ResultSheet = function () {


  var app = SpreadsheetApp.getActiveSpreadsheet();
  var gSheetName = 'Result';
  var gSheet = null;
  var gTitle = ['Result:', 'SparkLine:'];
  var gSparkHeader = ['Name', 'Set 1', 'Set 2', 'Set 3', 'Mixed', ' ']; //last column is empty to separate per week

  var lookupRange = null;
  var sparkLineRow = null;

  var util = new Util();

  var players = util.getPlayers();  //??? weird it already has guests ???


  


  this.initialize = function () {

      util.createSheet(gSheetName);

      //get created sheet at here
      gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(gSheetName);



  }

  this.addResultData = function () {

      //title    
      gSheet.getRange('A1').setValue(gTitle[0]).setFontWeight('bold');

      // result data
      gSheet.getRange('A2').setFormula('query(' + g.NamedRange_Month + ',"select * ")');

      //get location for lookup data for sparkline
      lookup_col_start = 'A'; //4
      lookup_row_start = gSheet.getLastRow() - players.length + 1; //4
      lookup_col_end = columnToLetter(gSheet.getLastColumn()); //'U'
      lookup_row_end = gSheet.getLastRow();

      lookupRange = '$' + lookup_col_start + '$' + lookup_row_start + ':' + '$' + lookup_col_end + '$' + lookup_row_end;

  }

  this.addSparkLineData = function () {

      lastRow = gSheet.getLastRow() + 4;
      a1Cell = "A" + lastRow;

      //title    
      gSheet.getRange(a1Cell).setValue(gTitle[1]).setFontWeight('bold');

      //table header
      lastRow = gSheet.getLastRow() + 1;
      g.GameDateInfo = util.getGameDates(g.Month, g.Day, g.Year);

      for (var n = 1; n <= g.GameDateInfo.numOfWeek; n++) {

          //each week header
          gSheet.getRange(lastRow, 5 * n - 3, 1, 5).setValues([gSparkHeader.slice(1)]);  //except Name column at here


          //add sparkline data
          var indexForSparkLine = 5 * n - 3;
          for (i = indexForSparkLine; i <= indexForSparkLine + 3; i++) {

              gSheet.getRange(lastRow + 1, i, 1, 1).setFormula('iferror(value(vlookup($' + 'A' + (lastRow + 1) + ', ' + lookupRange + ', ' + (i + 1) + ',false)), 0)');

          }

      }

      //Name Colums
      gSheet.getRange(lastRow, 1).setValue(gSparkHeader[0]); //Name column add at here

      //add members
      startRow = gSheet.getLastRow();
      sparkLineRow = startRow;


      players.sort();

      var tMember = players.map(function (elem) { return [elem]; });

      gSheet.getRange(startRow, 1, tMember.length, 1).setValues(tMember);

      //Copy to, first spark line data to other rows
      gSheet.getRange(startRow, 2, 1, g.GameDateInfo.numOfWeek * 5).copyTo(gSheet.getRange(startRow + 1, 2, players.length - 1, g.GameDateInfo.numOfWeek * 5));

  }

  this.setColumnWidth = function (colWidth) {

      var maxCol = gSheet.getMaxColumns();
      gSheet.setColumnWidths(2, maxCol - 1, colWidth);

  }

  this.mergeCells = function () {

      for (var i = 1; i <= g.GameDateInfo.numOfWeek; i++) {
          gSheet.getRange(2, 5 * i - 3, 1, 5).merge();

      }
  }

  this.setCellCenterAndMiddle = function () {

      setCellCenterAndMiddle(gSheetName);

  }

  this.deleteRemainedRowsAndColumns = function () {

      deleteRemainedRowsAndColumns(gSheetName, 2 /*remainedRows */, 2 /*remainedCols*/);

  }

  this.setNamedRange = function () {


      var lastCol = gSheet.getLastColumn();

      app.setNamedRange(g.NamedRange_Result, gSheet.getRange(4, 1, players.length, lastCol));

      app.setNamedRange(g.NamedRange_SparkLine, gSheet.getRange(sparkLineRow, 1, players.length, lastCol));



  }

}



