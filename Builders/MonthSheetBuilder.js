function MonthSheetBuilder() {


    mSheet = new MonthSheet();

    mSheet.initialize();
    mSheet.setNumberOfCell();
    mSheet.createHeader();
    mSheet.addMember();
    mSheet.setColor();
    mSheet.setInitialValue();
    mSheet.setDataValidation();
    mSheet.setFont();
    mSheet.setAlignment();
    mSheet.setSizeOfCell();
    mSheet.setConditionalRule();
    mSheet.mergeCells();
    mSheet.setBorder();
    mSheet.setNamedRange();
    mSheet.setFreeze();

    //set cells as plain text. It's important to calculate Win Ratio(%Win) value
    mSheet.setPlainText();

}


MonthSheet = function () {


    var app = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = g.Month;
    var sheet = null;


    numOfColPerWeek = 9;

    header = ['Name', 'In/Out', 'Set 1', 'Set 2', 'Set 3', 'Mixed'];

    backgroundColor = '#e4e7ff';
    mixedColor = 'yellow';
    inoutColor = '#93C47D';

    fontSize = 14;
    fontWeight = 'bold';
    cellWidth = 80;
    cellHeight = 80;
    outRowColor = '#9b7df6';
    frozenRow = 2;
    frozenCol = 1;

    util = new Util();

    var players = util.getPlayers();




    //methods
    this.initialize = function () {

        //util = new Util() ; 
        util.createSheet(sheetName);

        //get created sheet at here
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);


    };

    this.setNumberOfCell = function () {

        //remove rows
        var requiredRows = players.length + 2;
        if (sheet.getMaxRows() > requiredRows) {
            sheet.deleteRows(requiredRows, sheet.getMaxRows() - requiredRows);
        }
    };


    this.createHeader = function () {

        g.GameDateInfo = util.getGameDates(g.Month, g.Day, g.Year);

        // header
        for (var n = 1; n <= g.GameDateInfo.numOfWeek; n++) {

            //dates
            sheet.getRange(1, 5 * n - 3).setValue(g.GameDateInfo.dates[n - 1]);
            //each week header
            sheet.getRange(2, 5 * n - 3, 1, 5).setValues([header.slice(1)]);

        }

        //subHeader
        sheet.getRange(2, 1).setValue(header[0]); //Name

    }

    this.addMember = function () {

        for (var i = 0; i < players.length; i++) {
            sheet.getRange(3 + i, 1, 1, 1).setValue(players[i])
        }
    }

    this.setColor = function () {

        for (var i = 1; i <= Math.ceil(g.GameDateInfo.numOfWeek / 2); i++) {

            sheet.getRange(1, 5 * 2 * i - 8, players.length + 2, 5).setBackground(backgroundColor);

        }



        for (i = 1; i <= g.GameDateInfo.numOfWeek; i++) {

            //mixed color
            sheet.getRange(1, 5 * i + 1, players.length + 2, 1).setBackground(mixedColor);


        }


    }


    this.setInitialValue = function () {

        for (var n = 1; n <= g.GameDateInfo.numOfWeek; n++) {

            for (j = 0; j <= 4; j++) {

                sheet.getRange(3, 5 * n - 3 + j, players.length, 1).setValue('NA');
            }

        }

    }

    this.setDataValidation = function () {

        var ruleInOut = SpreadsheetApp.newDataValidation()
            .requireValueInList(['NA', 'In', 'Out'], true)
            .build();

        var ruleScore = SpreadsheetApp.newDataValidation()
            .requireValueInList(['0', '1', '2', '3', '4', '5', '6', 'NA'], true)
            .build();


        for (var n = 1; n <= g.GameDateInfo.numOfWeek; n++) {

            //In/Out
            sheet.getRange(3, 5 * n - 3, players.length, 1).setDataValidation(ruleInOut);

            for (j = 1; j <= 4; j++) {
                sheet.getRange(3, 5 * n - 3 + j, players.length, 1).setDataValidation(ruleScore);
            }

        }

    }

    this.setFont = function () {

        var maxCol = sheet.getMaxColumns();
        var maxRow = sheet.getMaxRows();

        sheet.getRange(1, 1, maxRow, maxCol)
            .setFontSize(fontSize)
            .setFontWeight(fontWeight);

    }


    this.setAlignment = function () {

        var maxCol = sheet.getMaxColumns();
        var maxRow = sheet.getMaxRows();

        sheet.getRange(1, 1, maxRow, maxCol)
            .setVerticalAlignment('middle')
            .setHorizontalAlignment('center');


    }

    this.setSizeOfCell = function () {


        var maxCol = sheet.getMaxColumns();
        var maxRow = sheet.getMaxRows();

        sheet.setRowHeights(1, maxRow, cellHeight);
        sheet.setColumnWidths(1, maxCol, cellWidth);

        //special rowHeights
        sheet.setRowHeights(1, 1, 20);
        sheet.setRowHeights(2, 1, 20);

    }

    this.setConditionalRule = function () {

        var lastCol = sheet.getLastColumn();
        var lastRow = sheet.getLastRow();

        //In
        var range = sheet.getRange(1, 1, lastRow, lastCol);
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('In')
            .setFontColor('#FF0000')
            .setRanges([range])
            .build()
        var rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);


        //Out
        var range = sheet.getRange(1, 1, lastRow, lastCol);
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Out')
            .setFontColor('cyan')
            .setBackground(outRowColor)
            .setRanges([range])
            .build()
        var rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);

        //Win
        var range = sheet.getRange(1, 1, lastRow, lastCol);
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(6)
            .setFontColor('red')
            .setRanges([range])
            .build()
        var rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);

        //Lose
        var range = sheet.getRange(1, 1, lastRow, lastCol);
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(0, 5)
            .setFontColor('blue')
            .setRanges([range])
            .build()
        var rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);


        //Out Rows  
        for (var i = 1; i <= g.GameDateInfo.numOfWeek; i++) {

            var refCell = sheet.getRange(3, 5 * i - 3, 1, 1).getA1Notation();
            Logger.log(refCell);
            var range = sheet.getRange(3, 5 * i - 2, 1, 4);
            var rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=$' + refCell + '="Out"')
                .setFontColor('white')
                .setBackground(outRowColor)
                .setRanges([range])
                .build()
            var rules = sheet.getConditionalFormatRules();
            rules.push(rule);
            sheet.setConditionalFormatRules(rules);

            //copy to others
            sheet.getRange(3, 5 * i - 2, 1, 4).copyTo(sheet.getRange(4, 5 * i - 2, g.Members.length - 1, 4));
        }

    }

    this.mergeCells = function () {

        //first row merging
        for (var i = 1; i <= g.GameDateInfo.numOfWeek; i++) {
            sheet.getRange(1, 5 * i - 3, 1, 5).merge();
        }

    }

    this.setBorder = function () {

        for (var n = 1; n <= g.GameDateInfo.numOfWeek; n++) {

            sheet.getRange(2, 5 * n - 3, g.Members.length + 1, 1).setBorder(false, true, false, true, false, false);

        }

    }

    this.setFreeze = function () {

        sheet.setFrozenRows(frozenRow);
        sheet.setFrozenColumns(frozenCol);

    }

    this.setNamedRange = function () {

        var lastCol = sheet.getLastColumn();

        app.setNamedRange(g.NamedRange_Month, sheet.getRange(1, 1, g.Members.length + 2, lastCol))

    }

    this.setPlainText = function () {

        //get all cells
        var allCell = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

        // Plain text
        allCell.setNumberFormat("@");


    }



}