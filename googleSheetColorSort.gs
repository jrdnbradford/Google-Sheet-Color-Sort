/*
   Google Sheet-bound script that assists with sorting
   Google Sheet rows by background fill color

   License: MIT (c) 2020 Jordan Bradford
   GitHub: jrdnbradford

   Recommended OAuth Scope:
   https://www.googleapis.com/auth/spreadsheets.currentonly
*/


const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const Ui = SpreadsheetApp.getUi();
const appTitle = "Google Sheet Color Sort";
const alertHeaderMsg = "Does this Google Sheet have a header?";


function onOpen(e) {
    Ui.createMenu("Google Sheet Color Sort")
        .addItem("Sort Rows by Color", "showSortPrompt")
        .addItem("Add Sorting Column", "showColumnPrompt")
        .addToUi();

    activeSpreadsheet
        .toast("MIT (c) 2020 Jordan Bradford", "License", -1);
}


function showSortPrompt() {
    // Used to pass arguments to colorSort that automatically sort the entire Sheet
    const response = Ui.alert(appTitle, alertHeaderMsg, Ui.ButtonSet.YES_NO);
    const activeSheet = SpreadsheetApp.getActiveSheet();

    if (response == Ui.Button.YES) {
        colorSort(true, true, activeSheet);
    } else if (response == Ui.Button.NO) {
        colorSort(false, true, activeSheet);
    } else {
        activeSpreadsheet.toast("You clicked the close button.");
    }
}


function showColumnPrompt() {
    // Used to pass arguments to colorSort that add a sorting column for manual sorting
    const response = Ui.alert(appTitle, alertHeaderMsg, Ui.ButtonSet.YES_NO);
    const activeSheet = SpreadsheetApp.getActiveSheet();

    if (response == Ui.Button.YES) {
        colorSort(true, false, activeSheet);
    } else if (response == Ui.Button.NO) {
        colorSort(false, false, activeSheet);
    } else {
        activeSpreadsheet.toast("You clicked the close button.");
    }
}


function colorSort(hasHeader, sortRows, sheet) {
    const lastRow = sheet.getLastRow();
    const sortColumn = sheet.getLastColumn() + 1;
    sheet.hideColumns(sortColumn);

    let startRow;
    let numRows;
    if (hasHeader) {
        startRow = 2;
        numRows = lastRow - 1;
        sheet.setFrozenRows(1);
        if (!sortRows) {
            sheet.getRange(1, sortColumn).setValue("Sort Column");
        }
    } else {
        startRow = 1;
        numRows = lastRow;
    }

    // The background fill color of first cell in each row of the first
    // column is used to identify the entire row's background fill color
    let hexRange = sheet.getRange(startRow, sortColumn, numRows);
    let backgrounds = sheet.getRange(startRow, 1, numRows).getBackgrounds();
    hexRange.setValues(backgrounds);

    if (sortRows) {
        sheet.sort(sortColumn);
        hexRange.clear({contentsOnly: true});
    }

    sheet.showColumns(sortColumn);
}