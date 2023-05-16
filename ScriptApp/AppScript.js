
function makeTabs() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheets = ss.getSheets();

    var last = sheet.getLastRow();//identifies the last active row on the sheet

    var existed = true;

    //loop through the code until each row creates a tab.
    for (var i = 1; i < last; i++) {

        //get the range in column A and get the value.
        var tabName = sheet.getRange(i + 1, 1).getValue() + " " + sheet.getRange(i + 1, 2).getValue();

        var amSheet = tabName + " AM";
        //create a new sheet with value for AM

        existed = existedTab(amSheet);


        if (existed == false) {
            ss.insertSheet(amSheet);

        }
        //get current AM sheet

        //count record without header

        var formulaBuilderAM = "=COUNTA(\'" + amSheet + "\'!A2:A)";
        //assign cell to formula
        sheet.getRange(i + 1, 5).setValue(formulaBuilderAM);
        Logger.log(formulaBuilderAM)



        var pmSheet = tabName + " PM";
        existed = existedTab(pmSheet);

        if (existed == false) {
            ss.insertSheet(pmSheet);

        }


        //get current PM sheet

        var formulaBuilderPM = "=COUNTA(\'" + pmSheet + "\'!A2:A)";
        //assign cell to formula
        sheet.getRange(i + 1, 6).setValue(formulaBuilderPM);
        Logger.log(formulaBuilderPM)


    }

}

function existedTab(tabToCompare) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheets = ss.getSheets();
    var existed = true;
    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetName().toLowerCase() == tabToCompare.toLowerCase()) {
            existed = true;
            break;
        }
        else {
            existed = false;
        }

    }
    return existed;
}

function deleteTab() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    Logger.log(sheets);
    for (var i = 1; i < sheets.length; i++) {
        ss.deleteSheet(sheets[i]);
        Logger.log(sheets[i]);
    }

}