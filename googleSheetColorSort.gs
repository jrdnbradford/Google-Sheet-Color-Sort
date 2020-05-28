/*
   Google Sheet-bound script that assists with sorting
   Google Sheet rows by selected column's background fill color.

   License: MIT (c) 2020 Jordan Bradford
   GitHub: jrdnbradford

   Recommended OAuth Scope:
   https://www.googleapis.com/auth/spreadsheets.currentonly
*/


const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const Ui = SpreadsheetApp.getUi();
const appTitle = "Google Sheet Color Sort";
const alertHeaderMsg = "Does this Google Sheet have a header?";


function onOpen() {
    Ui.createMenu("Google Sheet Color Sort")
        .addItem("Sort Rows by Color", "sortRowByColorPrompt")
        .addItem("Add Sorting Column", "addSortingColumnPrompt")
        .addToUi();

    activeSpreadsheet
        .toast("MIT (c) 2020 Jordan Bradford", "License", -1);
}


function sortRowByColorPrompt() {
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


function addSortingColumnPrompt() {
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
    const selectedColumn = sheet.getCurrentCell().getColumn();
    const sortColumn = sheet.getLastColumn() + 1;

    if (sortRows) { // Hide use of hex color codes
        sheet.hideColumns(sortColumn);
    }

    let startRow;
    let numRows;
    if (hasHeader) {
        startRow = 2;
        numRows = lastRow - 1;
        sheet.setFrozenRows(1);
    } else { // No header
        startRow = 1;
        numRows = lastRow;
    }

    // Hex color codes of data range in first column
    const columnRangeBackgrounds = sheet.getRange(startRow, selectedColumn, numRows).getBackgrounds();
    // Column range in which to place hex color codes for sorting
    const hexColumnRange = sheet.getRange(startRow, sortColumn, numRows);
    hexColumnRange.setValues(columnRangeBackgrounds);

    if (sortRows) {
        sheet.sort(sortColumn);
        hexColumnRange.clear({contentsOnly: true});
        sheet.showColumns(sortColumn);
    } else {
        if (hasHeader) {
            sheet.getRange(1, sortColumn).setValue("Sort Column");
        }
    }
}