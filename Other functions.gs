/**
 * PRINTING AN ARRAY TO A SPECIFIC SHEET, ROW & COL
 * @param {string} sheetName: Sheet where you want the Array printed
 * @param {Array<string>} arrayToPrint: Array you want to print
 * @param {number} rowToPrint: Row number where you want the array printed
 * @param {number} colToPrint: Column number where you want the array printed.
 */

function printTo_(sheetName,arrayToPrint,rowToPrint,colToPrint) {
    let ss = SpreadsheetApp.getActiveSpreadsheet()
    let sheetToPrint = ss.getSheetByName(sheetName);
    sheetToPrint.getRange(rowToPrint, colToPrint, arrayToPrint.length, arrayToPrint[0].length).setValues(arrayToPrint);
}

