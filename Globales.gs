// Source: https://www.youtube.com/watch?v=vdP6sZKp4hU

/**
 * @fileOverview This file replicates the Vlookup function between arrays.
 * @author Joaquin Pagliettini (www.hoakeen.com)
 * @version 1.0
 */

// Globals
let hojaSource = 'P1'      // [<EDIT HERE>]
let hojaSearch = 'P2' // [<EDIT HERE>]



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

