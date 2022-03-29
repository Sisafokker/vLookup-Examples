// Vlookup for a Range of strings

function vlookupEntireColumn() {
  let printArray = [['Sede', 'ClassID', 'Google searchroom Name']];  // [<EDIT HERE>]
//  let reorderArray = [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Source Sheet
  const sheetSource = ss.getSheetByName(hojaSource);
  const getSourceColRange = [[1], [5]];
  const getSourceLastRow = sheetSource.getLastRow();
  const rangeSource = sheetSource.getRange(2, getSourceColRange[0], getSourceLastRow, getSourceColRange[1])
  const sourceData = rangeSource.getValues();

  // Searching Sheet
  const sheetSearch = ss.getSheetByName(hojaSearch);
  const rangeSearch = sheetSearch.getDataRange();
  const allDataSearch = rangeSearch.getValues();
  // console.log({sourceData})
  // console.log({allDataSearch});

  // Search Column I want to match
  const searchInColumn = 1
  const searchLastRow = sheetSearch.getLastRow();
  const searchRange = sheetSearch.getRange(1, searchInColumn, searchLastRow);
  const searchData = searchRange.getValues();
  //  console.log({searchData});

  // Searching for each item in this ARRAY
  const searchArray = sheetSource.getRange(2, 1, getSourceLastRow).getValues();
  //console.log(searchArray);
  for (i = 1; i < searchArray.length; i++) {
    var searchString = searchArray[i];
    if (searchString == '') searchString = "Nothing"

    // We apply the prototpe finder2 generated below to the Array that has should have the match
    // Positive Match: Returns the row number if found. ;  Negativa Match: 
    // console.log({searchString});
    const arrayRowMatch = searchData.finder2(searchString);
    const sheetRowMatch = arrayRowMatch + 1
    console.log("arrayRowMatch %s / SheetRowMatch %s", arrayRowMatch, sheetRowMatch);

    if (arrayRowMatch && arrayRowMatch > -1) {
      var arrayMatch = allDataSearch[arrayRowMatch]
      printArray.push([arrayMatch[1], arrayMatch[4], arrayMatch[3]]);
    } else {
      printArray.push(['', '', '']);
    }
    Logger.log(printArray);
  }
  // Printing to Sheet.
  printTo_(hojaSource,printArray,2,5)
}

// Javascript search prototype for an Array? (dig into this...))
Array.prototype.finder2 = function (valor) {
  if (valor == "") return false;
  for (let i = 0; i < this.length; i++) {                     // this Refers to the Array
    if (this[i].toString().indexOf(valor) > -1) return i;   // i would be the row the content was found in
  }
  Logger.log("Not Found")
  return -1;  // if not found
};

