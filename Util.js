Util = function () {

    //create a new sheet
    this.createSheet = function (sheetName) {

        var app = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = app.getSheetByName(sheetName);

        if (sheet == null) {
            app.insertSheet(sheetName);
        }
        else {

            //delete an old sheet and create new one.
            var sheets = app.getSheets();

            if (sheets.length == 1) {

                //rename it
                app.renameActiveSheet("ToBeRemoved");

                // insert new sheet
                app.insertSheet(sheetName);

                //delete an old sheet
                app.deleteSheet(app.getSheetByName("ToBeRemoved"));

            }
            else {
                app.deleteSheet(sheet);
                app.insertSheet(sheetName);
            }
        }


    }


    this.setSize = function (startRow, endRow, startCol, endCol, width, height) {

        if (width != 0) {
            this.sheet.setColumnWidths(startCol, endCol - startCol + 1, width);
        }

        if (height != 0) {
            this.sheet.setRowHeights(startRow, endRow - startRow + 1, height);
        }

    };

    this.getGameDates = function (monthName, dayName, year) {

        // set names
        var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        var dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thr', 'Fri', 'Sat'];

        // change string to index of array
        var day = dayNames.indexOf(dayName);
        var month = monthNames.indexOf(monthName) + 1;

        // determine the number of days in month
        var daysinMonth = new Date(year, month, 0).getDate();

        // set counter
        var numOfWeek = 0;
        var gameDates = [];

        // iterate over the days and compare to day
        for (var i = 1; i <= daysinMonth; i++) {

            var targetDate = new Date(year, month - 1, parseInt(i));
            var dateInfo = Utilities.formatDate(targetDate, 'GMT', 'MM/dd/yyyy')

            var checkDay = targetDate.getDay();

            if (day == checkDay) {
                numOfWeek++;
                gameDates.push(dateInfo);

            }
        }


        return { numOfWeek: numOfWeek, dates: gameDates };

    }


    this.unfreezeAll = function () {

        var app = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = app.getSheets()[sheetIndex];

        SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetIndex].setFrozenColumns(0);
        SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetIndex].setFrozenRows(0);

    }

    this.getPlayers = function () {


        var players = g.Members;


        //Guest 
        if (g.Guests != null && g.Guests.length > 0) {

            for (i = 0; i < g.Guests.length; i++) {
                players.push(g.Guests[i]);
            }

            g.Guests = null;  // reset after add it ???, I do not know why g.Members are already have this guest data after one loop.
        }

        return players.sort();

    }

}




function columnToLetter(column) {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

function letterToColumn(letter) {
    var column = 0, length = letter.length;
    for (var i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
}


function setCellCenterAndMiddle(sheetName) {

    var app = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = app.getSheetByName(sheetName);

    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();

    sheet.getRange(1, 1, maxRow, maxCol)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');

}


function deleteRemainedRowsAndColumns(sheetName, remainedRows, remainedColumns) {

    var app = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = app.getSheetByName(sheetName);

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    var maxRow = sheet.getMaxRows();
    var maxCol = sheet.getMaxColumns();


    sheet.deleteRows(lastRow + 1, maxRow - lastRow - remainedRows);  //(rowPosition, howMany)


    if (maxCol - lastCol >= remainedColumns) {

        sheet.deleteColumns(lastCol + 1, maxCol - lastCol - remainedColumns);
    }


}

function setColumnWidth(sheetName, colWidth) {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    var maxCol = sheet.getMaxColumns();
    sheet.setColumnWidths(1, maxCol, colWidth);

}






