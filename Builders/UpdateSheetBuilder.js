function UpdateSheetBuilder() {

    updateSheet = new UpdateSheet();

    updateSheet.initialize();
    updateSheet.addWeekSelector();
    updateSheet.setDataValidation();
    updateSheet.setQueryData();

    updateSheet.addUpdateData();

    updateSheet.setNamedRange();

    updateSheet.deleteRemainedRowsAndColumns();
    updateSheet.setCellCenterAndMiddle();


}

UpdateSheet = function () {


    var app = SpreadsheetApp.getActiveSpreadsheet();
    var gSheetName = "Update";
    var gSheet = null;
    var gTitle = ['Week:', 'Query:', 'Update:'];

    var weekHeader = ['Parameters', 'Start Week', 'End Week'];
    var queryHeader = ['Parameters', 'Subject Columns', 'Target Columns'];

    var util = new Util();

    var startRow = 1;


    GameDateInfo = util.getGameDates(g.Month, g.Day, g.Year);


    var query_InOut = (GameDateInfo.numOfWeek > 4) ? "B,G,L,Q,V" : "B,G,L,Q";
    var query_Set = (GameDateInfo.numOfWeek > 4) ? "C,D,E,F,H,I,J,K,M,N,O,P,R,S,T,U,W,X,Y,Z" : "C,D,E,F,H,I,J,K,M,N,O,P,R,S,T,U";
    var query_WinSet = (GameDateInfo.numOfWeek > 4) ? "C,D,E,H,I,J,M,N,O,R,S,T,W,X,Y" : "C,D,E,H,I,J,M,N,O,R,S,T";
    var query_MinWinSet = (GameDateInfo.numOfWeek > 4) ? "F,K,P,U,Z" : "F,K,P,U";


    var header = [['Group\nSeed', 'Name', '#In', '#Out', '#Set', '#Win\nSet', '#Win\nMixed', '%Win\nSet', 'Credit', '#Win\nGame', 'Last M\nCredit']];

    var players = util.getPlayers();

    this.initialize = function () {


        util.createSheet(gSheetName);

        //get created sheet at here
        gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(gSheetName);



    }

    this.addWeekSelector = function () {

        gSheet.getRange(startRow, 1).setValue(gTitle[0]).setFontWeight('Bold');

        gSheet.getRange(startRow + 1, 1, 1, weekHeader.length).setValues([weekHeader]).setFontWeight('Bold')

        gSheet.getRange(startRow + 2, 1).setValue('Select Weeks').setFontWeight('Bold');

        gSheet.getRange(startRow + 1, 1, 2, 3).setBorder(true, true, true, true, true, true);

    }

    this.setDataValidation = function () {

        var weeks = ['Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5'];

        if (GameDateInfo.numOfWeek < 5) {

            weeks.pop();  // remove wk 5

        }

        var dropDown_Week = SpreadsheetApp.newDataValidation()
            .requireValueInList(weeks, true)
            .build();


        gSheet.getRange(startRow + 2, 2, 1, 2).setDataValidation(dropDown_Week).setBackground('yellow');

        //initial values
        gSheet.getRange(startRow + 2, 2, 1, 1).setValue(weeks[1]);
        gSheet.getRange(startRow + 2, 3, 1, 1).setValue(weeks[2]);




    }

    this.setQueryData = function () {


        startRow = 6;

        gSheet.getRange(startRow, 1).setValue(gTitle[1]).setFontWeight('Bold');

        //row merging
        gSheet.getRange(startRow + 1, 2, 1, 3).merge();
        gSheet.getRange(startRow + 1, 5, 1, 3).merge();

        //copy merged cells
        gSheet.getRange(startRow + 1, 1, 1, 7).copyTo(gSheet.getRange(startRow + 2, 1, 5, 7));


        gSheet.getRange(startRow + 1, 1).setValue(queryHeader[0]).setFontWeight('Bold');
        gSheet.getRange(startRow + 1, 2).setValue(queryHeader[1]).setFontWeight('Bold');
        gSheet.getRange(startRow + 1, 5).setValue(queryHeader[2]).setFontWeight('Bold');

        //subject columns
        gSheet.getRange(startRow + 2, 1).setValue('In/Out').setFontWeight('Bold');
        gSheet.getRange(startRow + 2, 2).setValue(query_InOut);

        gSheet.getRange(startRow + 3, 1).setValue('Set').setFontWeight('Bold');
        gSheet.getRange(startRow + 3, 2).setValue(query_Set);

        gSheet.getRange(startRow + 4, 1).setValue('#Win Set').setFontWeight('Bold');
        gSheet.getRange(startRow + 4, 2).setValue(query_WinSet);

        gSheet.getRange(startRow + 5, 1).setValue('#Win Set(Mix)').setFontWeight('Bold');
        gSheet.getRange(startRow + 5, 2).setValue(query_MinWinSet);

        // target columns
        gSheet.getRange(startRow + 2, 5).setFormula('=mid($B$8, 2*right($B$3)-1, (2*right($C$3)-1) - (2*right($B$3)-1)  + 1)');
        gSheet.getRange(startRow + 3, 5).setFormula('=mid($B$9, 8*right($B$3)-7, (8*right($C$3)-7) - (8*right($B$3)-7)  + 7)');
        gSheet.getRange(startRow + 4, 5).setFormula('=mid($B$10, 6*right($B$3)-5, (6*right($C$3)-5) - (6*right($B$3)-5)  + 5)');
        gSheet.getRange(startRow + 5, 5).setFormula('=mid($B$11, 2*right($B$3)-1, (2*right($C$3)-1) - (2*right($B$3)-1)  + 1)');

        gSheet.getRange(startRow + 1, 1, 5, 7).setBorder(true, true, true, true, true, true);


    }

    this.addUpdateData = function () {

        startRow = 14;

        gSheet.getRange(startRow, 1).setValue(gTitle[2]).setFontWeight('Bold');

        gSheet.getRange(startRow + 1, 1, 1, header[0].length).setValues(header).setFontWeight('bold');


        //group number
        for (i = 1; i <= 4; i++) {

            gSheet.getRange((startRow + 2) + 4 * i - 4, 1, 4, 1).setValue(i);

        }






        //add members
        players.sort();
        var tMember = players.map(function (elem) { return [elem]; });
        gSheet.getRange(startRow + 2, 2, tMember.length, 1).setValues(tMember).setBackground('yellow').setFontWeight('bold');


        //In
        gSheet.getRange(startRow + 2, 3, tMember.length, 1).setFormula('countif(query(' + g.NamedRange_Result + ', concatenate(\"select \", $E$8, \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),\"In\")');

        //Out
        gSheet.getRange(startRow + 2, 4, tMember.length, 1).setFormula('countif(query(' + g.NamedRange_Result + ', concatenate(\"select \", $E$8, \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),\"Out\")');

        //Set
        gSheet.getRange(startRow + 2, 5, tMember.length, 1).setFormula('countif(query(' + g.NamedRange_Result + ', concatenate(\"select \", $E$9, \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),\"<>NA\")');
        //Win Set
        gSheet.getRange(startRow + 2, 6, tMember.length, 1).setFormula('countif(query(' + g.NamedRange_Result + ', concatenate(\"select \", $E$10, \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),6)');

        //Win Mixed
        gSheet.getRange(startRow + 2, 7, tMember.length, 1).setFormula('countif(query(' + g.NamedRange_Result + ', concatenate(\"select \", $E$11, \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),6)');

        //%Win Set
        gSheet.getRange(startRow + 2, 8, tMember.length, 1).setFormula('if($E' + (startRow + 2) + '>0, ($F' + (startRow + 2) + ' + $G' + (startRow + 2) + ' )/$E' + (startRow + 2) + ',0)').setNumberFormat("0.0%");

        //Credit
        gSheet.getRange(startRow + 2, 9, tMember.length, 1).setFormula('C' + (startRow + 2) + ' + F' + (startRow + 2) + ' + (G' + (startRow + 2) + ' *0.5)');



        //#Win Game
        gSheet.getRange(startRow + 2, 10, tMember.length, 1).setFormula('SUMPRODUCT(query(' + g.NamedRange_Result + ', concatenate(\"select \", $E$9, \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")))');


        //Last M Credit

        var LastGameDateInfo = util.getGameDates(g.LastMonth, g.LastDay, g.LastYear);

        var Lquery_InOut = (LastGameDateInfo.numOfWeek > 4) ? "B,G,L,Q,V" : "B,G,L,Q";
        var Lquery_Set = (LastGameDateInfo.numOfWeek > 4) ? "C,D,E,F,H,I,J,K,M,N,O,P,R,S,T,U,W,X,Y,Z" : "C,D,E,F,H,I,J,K,M,N,O,P,R,S,T,U";
        var Lquery_WinSet = (LastGameDateInfo.numOfWeek > 4) ? "C,D,E,H,I,J,M,N,O,R,S,T,W,X,Y" : "C,D,E,H,I,J,M,N,O,R,S,T";
        var Lquery_MinWinSet = (LastGameDateInfo.numOfWeek > 4) ? "F,K,P,U,Z" : "F,K,P,U";

        var lastM_In = 'countif(query(' + g.NamedRange_LastMonth + ', concatenate(\"select \", \"' + Lquery_InOut + '\", \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),\"In\")';
        var lastM_WinSet = '+ countif(query(' + g.NamedRange_LastMonth + ', concatenate(\"select \", \"' + Lquery_WinSet + '\", \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),6) ';
        var lastM_WinMix = ' + (countif(query(' + g.NamedRange_LastMonth + ', concatenate(\"select \", \"' + Lquery_MinWinSet + '\", \" where A=\'\"&$B' + (startRow + 2) + '&\"\'\")),6) ) * 0.5';


        gSheet.getRange(startRow + 2, 11, tMember.length, 1).setFormula(lastM_In + lastM_WinSet + lastM_WinMix);


    }

    this.setNamedRange = function () {

        var lastCol = gSheet.getLastColumn();

        app.setNamedRange(g.NamedRange_Update, gSheet.getRange((startRow + 2), 1, g.Members.length, lastCol))

    }


    this.deleteRemainedRowsAndColumns = function () {

        deleteRemainedRowsAndColumns(gSheetName, 2 /*remainedRows */, 2 /*remainedCols*/);

    }

    this.setCellCenterAndMiddle = function () {

        setCellCenterAndMiddle(gSheetName);
    }



}