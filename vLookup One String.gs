// Vlookup for 1 string.

// Searching for...
const searchOneString = '22VirtualLEV 2-3 (12)SA-10:00/13:00 ONLINE';


function vlookupOneString() {
  let printArray =[['Sede','searchID','Google searchroom Name']];
  let reorderArray=[];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
// Source Sheet
  const sheetSource = ss.getSheetByName(hojaSource);
  const getSourceColRange = [[1] ,[5]];
    // console.log({getSourceCol[0]});
    // console.log({getSourceCol[1]});
  const getStRow = sheetSource.getLastRow();
  const rangeSource = sheetSource.getRange(2,getSourceColRange[0],getStRow,getSourceColRange[1])
//  const sourceData = rangeSource.getValues();

// Searching Sheet
  const sheetSearch = ss.getSheetByName(hojaSearch);
  const rangeSearch = sheetSearch.getDataRange();
  const allDataSearch = rangeSearch.getValues();
  // console.log({sourceData})
   //console.log({allDataSearch});

// search Column I want to match
  const searchInColumn = 1
  const searchLastRow = sheetSearch.getLastRow();
  const searchRange = sheetSearch.getRange(1,searchInColumn,searchLastRow);
  const searchData = searchRange.getValues();
  //  console.log({searchData});


// We apply the prototpe finder generated below to the Array that has should have the match
// Positive Match: Returns the row number if found. ;  Negativa Match: 
  const arrayRowMatch = searchData.finder1(searchOneString);
  const sheetRowMatch = arrayRowMatch+1
  console.log("arrayRowMatch %s / SheetRowMatch %s",arrayRowMatch, sheetRowMatch);

  
  if(arrayRowMatch > -1){
    var arrayMatch = allDataSearch[arrayRowMatch]
    reorderArray.push(arrayMatch[1],arrayMatch[4], arrayMatch[3]);
    printArray.push(reorderArray);
    Logger.log(printArray);
  }
// allDataSearch[arrayRowMatch] => [ [ '22VirtualLEV 2-3 (12)SA-10:00/13:00 ONLINE','Virtual', 'LEV 2-3 (12)', 'Level 2-3 SA.10a13.ONLINE (Virtual) [22]', 467794845063, 'x' ]
// printArray =>               [ [ 'Virtual', 467794845063, 'Level 2-3 SA.10a13.ONLINE (Virtual) [22]' ] ]

}

// Javascript search prototype for an Array? (dig into this...))
Array.prototype.finder1 = function (valor){
  if(valor =="") return false;
    for (let i=0; i< this.length ;i++){                     // this Refers to the Array
      if(this[i].toString().indexOf(valor) > -1) return i;   // i would be the row the content was found in
  } 
  return -1;
};