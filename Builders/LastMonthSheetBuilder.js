function LastMonthSheetBuilder() {

    lastMonthSheet = new LastMonthSheet();

    lastMonthSheet.initialize();
    lastMonthSheet.importLastMonthData();

    lastMonthSheet.deleteRemainedRowsAndColumns()
    lastMonthSheet.setCellCenterAndMiddle();

    lastMonthSheet.mergeCells(); //merge date cells

    lastMonthSheet.setColumnWidth(55); //column width

}

LastMonthSheet = function () {

    var app = SpreadsheetApp.getActiveSpreadsheet();
    var gSheetName = "Last M";
    var gSheet = null;  // there is no active sheet at this time. create a sheet first and then get a created sheet
    var gTitle = ['Last Month:'];
    var util = new Util();
    var GameDateInfo = util.getGameDates(g.Month, g.Day, g.Year);



    this.initialize = function () {


        util.createSheet(gSheetName);

        //get created sheet at here
        gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(gSheetName);


    }

    this.importLastMonthData = function () {


        //title    
        gSheet.getRange('A1').setValue(gTitle[0]).setFontWeight('bold');

        // get last month data
        gSheet.getRange('A2').setFormula('ImportRange(\"' + g.LastMonthSheetID + '\", \"' + g.NamedRange_Month + '\" )');


        //namedRange
        var lastCol = gSheet.getLastColumn();
        var lastRow = gSheet.getLastRow();
        app.setNamedRange(g.NamedRange_LastMonth, gSheet.getRange(4, 1, lastRow - 4 + 1, lastCol));


    }

    this.deleteRemainedRowsAndColumns = function () {

        deleteRemainedRowsAndColumns(gSheetName, 2 /*remainedRows */, 2 /*remainedCols*/);

    }

    this.setCellCenterAndMiddle = function () {

        setCellCenterAndMiddle(gSheetName);

    }

    this.mergeCells = function () {

        for (var i = 1; i <= GameDateInfo.numOfWeek; i++) {
            gSheet.getRange(2, 5 * i - 3, 1, 5).merge();

        }
    }

    this.setColumnWidth = function (colWidth) {

        var maxCol = gSheet.getMaxColumns();
        gSheet.setColumnWidths(2, maxCol - 1, colWidth);

    }



}