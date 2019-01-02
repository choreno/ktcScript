
function SeedSheetBuilder() {

    seed = new SeedSheet();

    seed.initialize();
    seed.createHeader();
    seed.setGroupColor();
    seed.setSize();


    seed.createSeed();
    seed.setFont();
    seed.setAlignment();

    seed.deleteRemainedRowsAndColumns();

}

SeedSheet = function () {

    //fields
    var app = SpreadsheetApp.getActiveSpreadsheet();
    var gSheetName = "Seed";
    var gSheet = null;

    var util = new Util();



    var GameDateInfo = util.getGameDates(g.Month, g.Day, g.Year);

    var header = [['Group\nSeed', 'Name', '#In', '#Out', '#Set', '#Win\nSet', '#Win\nMixed', '%Win\nSet', 'Credit', '#Win\nGame', 'Last M\nCredit', 'Current M\nResult']];

    var rowHeight = 40;
    var colWidth = 60;
    var boundaryRowHeight = 15;

    var fontSize = 12;
    var fontWeight = 'bold';

    var criteriaColor = '#D9D9D9';


    var players = util.getPlayers();




    //methods
    this.initialize = function () {

        util.createSheet(gSheetName);

        //get created sheet at here
        gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(gSheetName);


    }


    this.createHeader = function () {

        gSheet.getRange(1, 1, 1, header[0].length).setValues(header).setBackground('cyan');

    };


    this.setGroupColor = function () {



        var numberOfGroup = players.length / 4;

        //criteria color
        for (i = 1; i <= numberOfGroup; i++) { // until 3 groups
            switch (i) {
                case 1:
                    var groupColor = 'yellow';
                    break;
                case 2:
                    var groupColor = 'lightgreen';
                    break;
                case 3:
                    var groupColor = '#D5A6BD';
                    break;
            }
            gSheet.getRange(5 * i - 3, 1, 4, header[0].length).setBackground(groupColor);
            gSheet.getRange(5 * i - 3, gSheet.getLastColumn() - 3, 4, 4).setBackground(criteriaColor);
        }

        for (i = 1; i <= numberOfGroup; i++) {
            switch (i) {
                case 1:
                    var color = 'blue';
                    break;
                case 2:
                    var color = 'black';
                    break;
                case 3:
                    var color = 'red';
                    break;
            }
            gSheet.getRange(2, gSheet.getLastColumn() - (4 - i), players.length + 2, 1).setFontColor(color);
        }

    }

    this.setSize = function () {

        var maxCol = gSheet.getMaxColumns();
        var maxRow = players.length + 10;

        gSheet.setRowHeights(1, maxRow, rowHeight);
        gSheet.setColumnWidths(1, maxCol, colWidth);

        gSheet.setColumnWidth(gSheet.getLastColumn(), 140);

        //special rowHeights
        gSheet.setRowHeights(6, 1, boundaryRowHeight);
        gSheet.setRowHeights(11, 1, boundaryRowHeight);
        gSheet.setRowHeights(16, 1, boundaryRowHeight);

    }

    this.setFont = function () {

        var maxCol = gSheet.getLastColumn();
        var maxRow = gSheet.getLastRow();

        gSheet.getRange(1, 1, maxRow, maxCol)
            .setFontSize(fontSize)
            .setFontWeight(fontWeight);

    }


    this.setAlignment = function () {

        var maxCol = gSheet.getMaxColumns();
        var maxRow = gSheet.getMaxRows();

        gSheet.getRange(1, 1, maxRow, maxCol)
            .setVerticalAlignment('middle')
            .setHorizontalAlignment('center');


    }




    this.createSeed = function () {

        var startRow = 2;

        for (i = 1; i <= 3; i++) {

            //create seed by group
            gSheet.getRange(startRow + 5 * i - 5, 2).setFormula('query(' + g.NamedRange_Update + ',\"select B,C,D,E,F,G,H,I,J,K WHERE A=' + i + ' order by I desc, J desc, K desc, B\",-1)');

            gSheet.getRange(startRow + 5 * i - 5, 1).setValue(1);
            gSheet.getRange(startRow + 5 * i - 5 + 1, 1).setFormula('if($I3<$I2,$A2+1,if($J3<$J2,$A2+1,if($K3<$K2,$A2+1,$A2)))')

            //copy to
            gSheet.getRange(startRow + 5 * i - 5 + 1, 1).copyTo(gSheet.getRange(startRow + 5 * i - 5 + 2, 1, 2, 1));

        }


        //spark line

        var index = (GameDateInfo.numOfWeek > 4) ? "{2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25}" : "{2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20}";
        var sparkLineData = 'SPARKLINE(ArrayFormula(VLOOKUP($B2,' + g.NamedRange_SparkLine + ',' + index + ',FALSE)),{"charttype","column"; "ymin",0;"ymax",6;"highcolor","red";"empty","zero";"nan","convert"})';
        //var sparkLineData = 'SPARKLINE(ArrayFormula(VLOOKUP($B2,' + g.NamedRange_SparkLine + ',' + index + ',FALSE)),{"charttype","column"; "color","blue";"ymin",0;"ymax",6;"highcolor","red";"empty","zero";"nan","convert"})';
        gSheet.getRange(startRow, 12).setFormula(sparkLineData);

        //copy to
        gSheet.getRange(startRow, 12).copyTo(gSheet.getRange(startRow + 1, 12, 3, 1));
        gSheet.getRange(startRow, 12).copyTo(gSheet.getRange(startRow + 5, 12, 4, 1));
        gSheet.getRange(startRow, 12).copyTo(gSheet.getRange(startRow + 10, 12, 4, 1));



    }



    this.deleteRemainedRowsAndColumns = function () {

        deleteRemainedRowsAndColumns(gSheetName, 2 /*remainedRows */, 2 /*remainedCols*/)

    }





}