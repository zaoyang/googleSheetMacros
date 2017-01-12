

var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

/** 
 * Indent the text using CONCAT. This is similar to an Excel function but it's not 
 * available in Google Spreadsheet. 
 */ 
function indentText() {
  var values = spreadSheet.getActiveRange().getValues();
  var newValues = new Array();
  
  
  
  for (i = 0; i < values.length; i++) {
    if (values[i][0] != '') {
      newValues.push(['=CONCAT(REPT( CHAR( 160 ), 5),"' + values[i][0] + '")']);
    } else {
      newValues.push(['']);  
    }
  }
    
  spreadSheet.getActiveRange().setValues(newValues);
};



/** 
 * Remove all the tabs 
 */ 
function removeIndent() {
  var range = spreadSheet.getActiveRange()
  var displayValues = range.getDisplayValues()
  range.clearContent()

  var newValues = Array();   
  for(i = 0; i < displayValues.length; i++){
    displayValueStr = displayValues[i][0].toString()
    newValues.push([displayValueStr.trim()])
  } 
  
  spreadSheet.getActiveRange().setValues(newValues);
   
};

/** 
 * Menu Indent text, remove indent. 
 * TODO Reverse Indent Text. Don't have time to do it right now. 
 */ 
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = []

  menuEntries.push({
    name : "Indent Text",
    functionName : "indentText", 
  })  
  
  menuEntries.push({
    name : "Remove Indent", 
    functionName : "removeIndent", 
  });
    
  sheet.addMenu("Indenting Functions", menuEntries);

};
