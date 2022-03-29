// Vlookup Library
// VLOOKUP([sourceSheet, sourceRange], [searchSheet; searchInThisRange], [ResCol#1, ResCol#2, ResCol#3])
// VLOOKUP([Hoja1, A2:A]];[Hoja1, A2:E26]; [2, 5, 4])

/** 
 * Replicating VLookup function for Arrays
 * @param {Array<strings>} fromSpecsArray [sourceSheet, sourceRange]
 * @param {Array<strings>} searchInSpecsArray [searchSheet; searchInThisRange]
 * @param {Array<number>} returnColSpecsArray [Response Col#1, Response Col#2, Response Col#3]
 * @author Joaquin Pagliettini (www.hoakeen.com)
*/

// CUIDADO: 
// El HEADER de la columna de SOURCE tiene que coincidir exactamente con el HEADER de la columna de la hoja SEARCH

function runVLookupCursos(){
  const quienes = "ABMCursos";
  const deDonde = "CLASES21"
   const resultadoFinal = vlookupAndMix_([quienes,"A3:A","A3:H"], [1,2,7,8], [deDonde,"A3:A"], [4, 3, 5, 6])
   let [arrayCompleto, arraySoloResultado] = resultadoFinal
   console.log({arrayCompleto});
   console.log('============================================');
  // console.log({arraySoloResultado});
  // console.log('============================================');
   printTo_(quienes,arrayCompleto,3,11);
}

function runVLookupAlumnos(){
  const quienes = "ABMAlumnos";
  const deDonde = "CLASES21"
   const resultadoFinal = vlookupAndMix_([quienes, "A3:A","A3:L"], [1,2,8,9], [deDonde, "A3:A"], [4, 3, 5, 6])
   let [arrayCompleto, arraySoloResultado] = resultadoFinal
   console.log({arrayCompleto});
   //console.log({arraySoloResultado});
   console.log('======================');
   printTo_(quienes,arrayCompleto,3,14);
}

function runVLookupProfesores(){
  const quienes = "ABMProfesores";
  const deDonde = "CLASES21"
   const resultadoFinal = vlookupAndMix_([quienes, "A3:A","A3:I"], [1,2,9,8], [deDonde, "A3:A"], [4, 3, 5, 6])
   let [arrayCompleto, arraySoloResultado] = resultadoFinal
   console.log({arrayCompleto});
   console.log('============================================');
  // console.log({arraySoloResultado});
  // console.log('============================================');
   printTo_(quienes,arrayCompleto,3,11);
}


function vlookupAndMix_(fromSpecsArray, addToResult = {}, searchInSpecsArray, returnColSpecsArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 

// DESTRUCTURING ARGUMENTS/PARAMETERS
  let [inputSourcePage, inputRangeToVlookup, inputSourceFullRange] = fromSpecsArray;
  let [addToResultCol0, addToResultCol1, addToResultCol2, addToResultCol3] = addToResult;
  let [inputPageSearchIn, inputSearchInRange] = searchInSpecsArray;
  let [returnCol0, returnCol1, returnCol2, returnCol3] = returnColSpecsArray;
        // console.log({inputSourcePage, inputRangeToVlookup});
        // console.log({addToResultCol0, addToResultCol1, addToResultCol2, addToResultCol3});
        // console.log({inputPageSearchIn, inputSearchInRange});
        // console.log({returnCol0, returnCol1, returnCol2, returnCol3});

// SOURCE SHEET (1ST SHEET)
  const sourceSheet = ss.getSheetByName(inputSourcePage);
  const allDataInSourceSheet = sourceSheet.getRange(inputSourceFullRange).getValues();
    
  const sourceRangeToFind = sourceSheet.getRange(inputRangeToVlookup.toString());
  const sourceDataToFind_Unfiltered = sourceRangeToFind.getValues();
  const sourceDataClean = sourceDataToFind_Unfiltered.filter((rowVal) => rowVal != '');
          // console.log(inputSourcePage,{sourceDataClean});

// SEARCHING SHEET (2ND SHEET)
  const searchSheet = ss.getSheetByName(inputPageSearchIn);
  const allDataInSearchSheet = searchSheet.getDataRange().getValues();
          // console.log(inputSourcePage,{sourceDataClean})
          // console.log({allDataInSearchSheet});

// THE TWO EMPTY ARRAYS WE NEED TO POPULATE
  let sourceArray = []; // To add selected columns from Source Sheet
  let printArray = [];  // To return found columns from Search Sheet
  
// POPULATING SOURCE ARRAY BY FILTERING allDataInSourceSheet AND THEN PUSHING SPECIFIED COLUMNS
   const filteredDataSource = allDataInSourceSheet.filter((col) => col[addToResultCol0] !="");
  // console.log(filteredDataSource) ;
   for (var j = 0; j < filteredDataSource.length; j++){
      let dataSourceParcial = filteredDataSource[j]
          // console.log({dataSourceParcial});
      sourceArray.push([dataSourceParcial[addToResultCol0-1],dataSourceParcial[addToResultCol1-1],dataSourceParcial[addToResultCol2-1],dataSourceParcial[addToResultCol3-1]])
    }
         //  console.log({sourceArray}) ;


// FINDING VALUES IN SEARCH SHEET
   const searchInThisRange = searchSheet.getDataRange(); 
   //const searchInThisRange = searchSheet.getRange(inputSearchInRange.toString());  // If I activate this one, there is a 3 or 4 row index difference
   const searchInsideThisData = searchInThisRange.getValues();
           console.log(inputPageSearchIn,{searchInsideThisData});
   
// SEARCHING FOR EACH ITEM 
   for(i = 0; i < sourceDataClean.length; i++) {
     var sourceString = sourceDataClean[i];
     if (sourceString == '') sourceString = "Nothing"

     // We apply the prototpe myFinderMethod generated below to the Array that has should have the match
     // Positive Match: Returns the row number if found. ;  
     const indexOfMatch = searchInsideThisData.myFinderMethod(sourceString);
            // console.log({indexOfMatch})
     // Different index if its an array or a row in Google Sheets
     const gSheetRowMatch = indexOfMatch + 1
            // console.log("indexOfMatch %s / gSheetRowMatch %s", indexOfMatch, gSheetRowMatch);

     if (indexOfMatch && indexOfMatch > -1) {
       var fullRowOfMatch = allDataInSearchSheet[indexOfMatch]
       printArray.push([fullRowOfMatch[returnCol0 - 1], fullRowOfMatch[returnCol1 - 1], fullRowOfMatch[returnCol2 - 1], fullRowOfMatch[returnCol3-1]]);
     } else {
        printArray.push(['', '', '','']);
     }
    //console.log({printArray});
  }
  
  // RETURN
  var mixArray = []
  console.log(inputSourcePage, sourceArray.length);
  for (y = 0; y < sourceArray.length ; y++) {
   // console.log({y})
     mixArray.push(sourceArray[y].concat(printArray[y]));
    }
  return [mixArray, printArray];
}

// Javascript search prototype for an Array? (dig into this...))
Array.prototype.myFinderMethod = function (valor) {
  if (valor == "") return false;
  for (let i = 0; i < this.length; i++) {                     // this Refers to the Array
    if (this[i].toString().indexOf(valor) > -1) return i;   // i would be the row the content was found in
  }
  console.log(valor,"Not Found by Array Prototype")
  return -1;  // if not found
};

