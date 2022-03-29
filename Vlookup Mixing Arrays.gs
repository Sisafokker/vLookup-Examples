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

// CUIDADO: 
// El HEADER de la columna de SOURCE tiene que coincidir exactamente con el HEADER de la columna de la hoja SEARCH

function runVLookupCursos(){
  const quienes = "ABMCursos";
  const deDonde = "CLASES21"
   const resultadoFinal = vlookupAndMix_([quienes,"A3:A","A3:H"], [1,2,7,8], [deDonde,"A3:A"], [4, 3, 5, 6])
   let [arrayCompleto, arraySoloResultado] = resultadoFinal
   console.log({arrayCompleto});
   console.log('============================================');
   console.log({arraySoloResultado});
   console.log('============================================');
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
   console.log({arraySoloResultado});
   console.log('============================================');
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
  const sheetSource = ss.getSheetByName(inputSourcePage);
  const allDataSource = sheetSource.getRange(inputSourceFullRange).getValues();
    
  const rangeSourceToFind = sheetSource.getRange(inputRangeToVlookup.toString());
  const sourceDataToFindUnfiltered = rangeSourceToFind.getValues();
  const sourceData = sourceDataToFindUnfiltered.filter((rowVal) => rowVal != '');
          // console.log({sourceData});

// SEARCHING SHEET (2ND SHEET)
  const sheetSearch = ss.getSheetByName(inputPageSearchIn);
  const rangeSearch = sheetSearch.getDataRange();
  const allDataSearch = rangeSearch.getValues();
          // console.log({sourceData})
          // console.log({allDataSearch});

// THE TWO ARRAYS WE NEED TO POPULATE
  let sourceArray = []; // To add selected columns from Source Sheet
  let printArray = [];  // To return found columns from Search Sheet
  
// POPULATING SOURCE ARRAY BY FILTERING ALLDATASOURCE AND THEN PUSHING SPECIFIED COLUMNS
   const filteredDataSource = allDataSource.filter((col) => col[1] !="");
  // console.log(filteredDataSource) ;
   for (var j = 0; j < filteredDataSource.length; j++){
      let dataSourceParcial = filteredDataSource[j]
          // console.log({dataSourceParcial});
      sourceArray.push([dataSourceParcial[addToResultCol0-1],dataSourceParcial[addToResultCol1-1],dataSourceParcial[addToResultCol2-1],dataSourceParcial[addToResultCol3-1]])
    }
         //  console.log({sourceArray}) ;


// SEARCH COLUMN WE WANT TO MATCH
   const searchRange = sheetSearch.getDataRange(); /// No deberia ser asi...-----------------------------------
   const searchData = searchRange.getValues();
          // console.log({searchData});
   
// SEARCHING FOR EACH ITEM IN THE SOURCE ARRAY TO SEARCH
   for(i = 0; i < sourceData.length; i++) {
     var sourceString = sourceData[i];
     if (sourceString == '') sourceString = "Nothing"

     // We apply the prototpe myFinderMethod generated below to the Array that has should have the match
     // Positive Match: Returns the row number if found. ;  
     const indexOfMatch = searchData.myFinderMethod(sourceString);
            // console.log({indexOfMatch})
     // Different index if its an array or a row in Google Sheets
     const gSheetRowMatch = indexOfMatch + 1
            // console.log("indexOfMatch %s / gSheetRowMatch %s", indexOfMatch, gSheetRowMatch);

     if (indexOfMatch && indexOfMatch > -1) {
       var fullRowOfMatch = allDataSearch[indexOfMatch]
       printArray.push([fullRowOfMatch[returnCol0 - 1], fullRowOfMatch[returnCol1 - 1], fullRowOfMatch[returnCol2 - 1], fullRowOfMatch[returnCol3-1]]);
     } else {
        printArray.push(['', '', '','']);
     }
    //console.log({printArray});
  }
  
  // RETURN
  var mixArray = []
  console.log(sourceArray.length);
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

