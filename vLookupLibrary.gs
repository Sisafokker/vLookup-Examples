// Vlookup Library
// VLOOKUP([sourceSheet, sourceRange], [searchSheet; searchRange], [ResCol#1, ResCol#2, ResCol#3])
// VLOOKUP([Hoja1, A2:A]];[Hoja1, A2:E26]; [E, D, B])

/** 
 * Replicating VLookup function for Arrays
 * @param {Array<strings>} fromSpecsArray [sourceSheet, sourceRange]
 * @param {Array<strings>} searchInSpecsArray [searchSheet; searchRange]
 * @param {Array<number>} returnColSpecsArray [Response Col#1, Response Col#2, Response Col#3]
*/

function run(){
  vlookup_()
   vlookup_(["P1", "A3:A13"], ["P2", "A3:A"], [2, 5, 4])
}

function vlookup_(fromSpecsArray, searchInSpecsArray, returnColSpecsArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();  
  // Source Sheet
  const sheetSource = ss.getSheetByName(fromSpecsArray[0]);
  //const getSourceColRange = fromSpecsArray[1].toString();
  const getSourceColRange = fromSpecsArray[1].toString();
  const getSourceLastRow = sheetSource.getLastRow();
  const rangeSource = sheetSource.getRange(getSourceColRange);
  const sourceData = rangeSource.getValues();

  // Searching Sheet
  const sheetSearch = ss.getSheetByName(searchInSpecsArray[0]);
  const rangeSearch = sheetSearch.getDataRange();
  const allDataSearch = rangeSearch.getValues();
  // console.log({sourceData})
  // console.log({allDataSearch});

  // search Column I want to match
  // const searchInColumn = searchInSpecsArray[1];
  //const searchLastRow = sheetSearch.getLastRow();
  const searchRange = sheetSearch.getDataRange();
  const searchData = searchRange.getValues();
   console.log({searchData});
  
  let printArray = [['Sede', 'ClassID', 'Google searchroom Name']];
  
  // Searching for each item in this ARRAY
  for(i = 0; i < sourceData.length; i++) {
    var sourceString = sourceData[i];
    if (sourceString == '') sourceString = "Nothing"

    // We apply the prototpe finder generated below to the Array that has should have the match
    // Positive Match: Returns the row number if found. ;  Negativa Match: 
      console.log({sourceString});
    const arrayRowMatch = searchData.finder(sourceString);
    const sheetRowMatch = arrayRowMatch + 1
    console.log("arrayRowMatch %s / SheetRowMatch %s", arrayRowMatch, sheetRowMatch);

    if (arrayRowMatch && arrayRowMatch > -1) {
      var arrayMatch = allDataSearch[arrayRowMatch]
      printArray.push([arrayMatch[returnColSpecsArray[0]-1], arrayMatch[returnColSpecsArray[1]-1], arrayMatch[returnColSpecsArray[2]-1]]);
    } else {
      printArray.push(['', '', '']);
    }
    console.log({printArray});
  }
  // Printing to Sheet.
  printTo_(fromSpecsArray[0],printArray,2,5)
}

// Javascript search prototype for an Array? (dig into this...))
Array.prototype.finder = function (valor) {
  if (valor == "") return false;
  for (let i = 0; i < this.length; i++) {                     // this Refers to the Array
    if (this[i].toString().indexOf(valor) > -1) return i;   // i would be the row the content was found in
  }
  Logger.log("Not Found")
  return -1;  // if not found
};
