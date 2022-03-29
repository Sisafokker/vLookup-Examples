// Vlookup Library
// VLOOKUP([sourceSheet, sourceRange], [searchSheet; searchRange], [ResCol#1, ResCol#2, ResCol#3])
// VLOOKUP([Hoja1, A2:A]];[Hoja1, A2:E26]; [2, 5, 4])

/** 
 * Replicating VLookup function for Arrays
 * @param {Array<strings>} fromSpecsArray [sourceSheet, sourceRange]
 * @param {Array<strings>} searchInSpecsArray [searchSheet; searchRange]
 * @param {Array<number>} returnColSpecsArray [Response Col#1, Response Col#2, Response Col#3]
 * @author Joaquin Pagliettini (www.hoakeen.com)
*/

function run(){
   const resultadoFinal = vlookup_(["P1", "A3:A13"], ["P2", "A3:A"], [2, 5, 4])
   console.log({resultadoFinal});
}

function vlookup_(fromSpecsArray, searchInSpecsArray, returnColSpecsArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  // Destructuring Function Arguments/Parameters
  let [inputPageSource, inputSourceRange] = fromSpecsArray;
  let [inputPageSearch, inputSearchRange] = searchInSpecsArray;
  let [returnCol0, returnCol1, returnCol2] = returnColSpecsArray;
  // console.log({inputPageSource, inputSourceRange});
  // console.log({inputPageSearch, inputSearchRange});
  // console.log({returnCol0, returnCol1, returnCol2});

  // Source Sheet
  const sheetSource = ss.getSheetByName(inputPageSource);
  //const getSourceColRange = inputSourceRange.toString();
  const getSourceColRange = inputSourceRange.toString();
  const getSourceLastRow = sheetSource.getLastRow();
  const rangeSource = sheetSource.getRange(getSourceColRange);
  const sourceData = rangeSource.getValues();

  // Searching Sheet
  const sheetSearch = ss.getSheetByName(inputPageSearch);
  const rangeSearch = sheetSearch.getDataRange();
  const allDataSearch = rangeSearch.getValues();
  // console.log({sourceData})
  // console.log({allDataSearch});

  // search Column I want to match
  // const searchInColumn = inputPageSearch;
  //const searchLastRow = sheetSearch.getLastRow();
  const searchRange = sheetSearch.getDataRange();
  const searchData = searchRange.getValues();
   console.log({searchData});
  
  let printArray = [];
  
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
      printArray.push([arrayMatch[returnCol0-1], arrayMatch[returnCol1-1], arrayMatch[returnCol2-1]]);
    } else {
      printArray.push(['', '', '']);
    }
    //console.log({printArray});
  }
  
  // RETURN
  return printArray
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

