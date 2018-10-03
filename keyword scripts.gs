

function fooWords(arr) {
    var a = [], b = [], prev;

    arr.sort();
    for ( var i = 0; i < arr.length; i++ ) {
        if ( arr[i] !== prev ) {
            a.push(arr[i]);
            b.push(1);
        } else {
            b[b.length-1]++;
        }
        prev = arr[i];
    }
    
    var newA = [];
    
    while(a.length) newA.push(a.splice(0,1));

    return newA;
}

function fooWordCounts(arr) {
    var a = [], b = [], prev;

    arr.sort();
    for ( var i = 0; i < arr.length; i++ ) {
        if ( arr[i] !== prev ) {
//            a.push(arr[i]);
            b.push(1);
        } else {
            b[b.length-1]++;
        }
        prev = arr[i];
    }
    
    var newB = [];
    
    while(b.length) newB.push(b.splice(0,1));

    return newB;
}





function getKeywords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Keywords");
  var sheets = ss.getSheets();
  
  sheet.getRange(9, 2, sheet.getLastRow()-9+1, 2).clear({contentsOnly: true});  // Clear previous values
  
  var arrWords = [];  // Store Words
  var arrCounts = [];  // Store counts
  var ignoreWords = ["in", "IN", "of", "OF", "A", "a"];  // Collection of ignored words
  
  for (var i=0; i<sheets.length; i++)
  {
    var getSheet = sheets[i];
    var sheetName = getSheet.getName();
    
    
    // Ignore sheets other than months
    if (sheetName == "Yearly" || sheetName == "Monthly" || sheetName == "Initials-Monthly" || sheetName == "Initials-Day" || sheetName == "Initials-Search"
       || sheetName == "ASIN" || sheetName == "Leaderboard" || sheetName == "Keywords" || sheetName == "CLOSED ASIN" || sheetName == "RP"
        || sheetName == "Intermediate leaderboard" || sheetName == "Cancelled Orders" || sheetName == "Template") { continue; }
    
    
    var targetRow = lookup("Product Title", getSheet, 2, 3, "row");
    
    var getTitles = getSheet.getRange(targetRow+1, 2, getSheet.getLastRow()-targetRow).getValues();
    
    
    // Ignore unwanted characters that doesn't count as word
    for (var j=0; j<getTitles.length; j++)
    {
      var title = getTitles[j].toString();
      if (title.indexOf(",") >= 0) {
        title = replaceAll(title, ",", "");
      }
      if (title.indexOf("(") >= 0) {
        title = replaceAll(title, "(", "");
      }
      if (title.indexOf(")") >= 0) {
        title = replaceAll(title, ")", "");
      }
      if (title.indexOf(" - ") >= 0) {
        title = replaceAll(title, " - ", " ");
      }
      if (title.indexOf(" -") >= 0) {
        title = replaceAll(title, " -", " ");
      }
      if (title.indexOf("&") >= 0) {
        title = replaceAll(title, "&", " ");
      }
      if (title.indexOf('"') >= 0) {
        title = replaceAll(title, '"', '');
      }
      if (title.indexOf("'") >= 0) {
        title = replaceAll(title, "'", "");
      }
      if (title.indexOf(":") >= 0) {
        title = replaceAll(title, ":", "");
      }
      if (title.indexOf("+") >= 0) {
        title = replaceAll(title, "+", " ");
      }
      if (title.indexOf(".") >= 0) {
        title = replaceAll(title, ".", " ");
      }
      if (title.indexOf("/") >= 0) {
        title = replaceAll(title, "/", " ");
      }
      
      var thisWords = title.split(" ");
      arrWords = arrWords.concat(thisWords);
    }
  }
  
  var nonEmptyWords = [];
  
  for (var k=0; k<arrWords.length; k++)  // Checks if there are empty words first
  {
    var word = arrWords[k].toString();
    if (word == "" || word.length < 4) { continue; }
    
      var isIgnored = 0;
      
      var patt = new RegExp(/^[0-9]+$/);  // Regex for numbers only
      var res = patt.test(word);  // Check if the word is digits only
      
            if (res == false) {  // If not digits only, check for ignoring words
            
                  for (var l=0; l<ignoreWords.length; l++)  // Checks if the word matches with any of the ignoring words
                  {
                    var ignoreWord = ignoreWords[l].toString();
                    if (word == ignoreWord) {
                      isIgnored = 1;
                      Logger.log(k+" - "+word);
                      break;
                    }
                  }
                  
                  if (isIgnored == 0) { nonEmptyWords.push(word); }  // Push the word if it's non-empty, non-ignoring and not digits only
              
            }
  }
  
  var singleWords = fooWords(nonEmptyWords);  // Convert to 2D array and remove duplicates
  arrCounts = fooWordCounts(nonEmptyWords);  // Total count numbers for each word
  
  var lr = 9;//sheet.getLastRow()+2;
  sheet.getRange(lr, 2, singleWords.length).setValues(singleWords);
  sheet.getRange(lr, 3, arrCounts.length).setValues(arrCounts);
}






function catagorizeWords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Keywords");
  var sheets = ss.getSheets();
  
//  var targetRow = lookup("10", sheet, 2, 3, "row");
  var lr = sheet.getLastRow()-7;
  var lc = sheet.getLastColumn()-1;
  
//  sheet.getRange(9, 2, sheet.getLastRow()-9+1, 2).clear({contentsOnly: true});  // Clear previous values
  sheet.getRange(9, 4,sheet.getLastRow()-9+1, sheet.getLastColumn()-4+1).clearContent();
  var rng = sheet.getRange(8, 2, lr, lc).getValues();
  
  var arrTitles = [];
  var arrCatWords = [];
  var arrCounts = [];
  
  for (var i=0; i<sheets.length; i++)
  {
          var getSheet = sheets[i];
          var sheetName = getSheet.getName();
          
          
                    // Ignore sheets other than months
                    if (sheetName == "Yearly" || sheetName == "Monthly" || sheetName == "Initials-Monthly" || sheetName == "Initials-Day" || sheetName == "Initials-Search"
                       || sheetName == "ASIN" || sheetName == "Leaderboard" || sheetName == "Keywords" || sheetName == "Copy of Keywords" || sheetName == "CLOSED ASIN" || sheetName == "RP"
                        || sheetName == "Intermediate leaderboard" || sheetName == "Cancelled Orders" || sheetName == "Template") { continue; }
                    
          
          var getTitles = getSheet.getRange(5, 2, getSheet.getLastRow()-5+1).getValues();
          
          
                // Ignore unwanted characters that doesn't count as word
                for (var j=0; j<getTitles.length; j++)
                {
                        var title = getTitles[j][0].toString();
                        arrTitles.push(title);
                }
  }
  
  for (var k=1; k<rng.length; k++)
  {
        var word = rng[k][0];
        if(word==""){continue;}
        if(rng[k][1]<10){continue;}
        var temp =myFilter(arrTitles, word);

          for (var l=2; l<rng[0].length; l++)
          {
                  var catagory = rng[0][l];
                  if(catagory==""){
                         continue;
                  }
                  
                  var temp2=myFilter(temp, catagory);
                  //if (catagory != "") { getCount = catCounts(arrTitles, word, catagory); }
                  rng[k][l] = temp2.length;
          }
    
  }
//  Logger.log(rng);
  sheet.getRange(8, 2, lr, lc).setValues(rng);
}




function myFilter(arrTitles, word)
{
    
    function doesInclude(value, index, array)
    {
        if(value.indexOf(word)>-1) { return true; }
        else { return false; }
    
    }
    
    
    var b = arrTitles.filter(doesInclude);
    return b;
  

}













function catCounts(arrTitles, word, catagory) {
  var count = 0;
  
  for (var i=0; i<arrTitles.length; i++)
  {
    var title = arrTitles[i].toString();
    
        if (title.indexOf(word) >= 0 && title.indexOf(catagory) >= 0) {
          count += 1;
        }
    
  }
  
  return count;
}




