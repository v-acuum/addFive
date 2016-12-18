function modifyCell(qualityOfCompletion) {
  
  //Get the Spreadsheet, Sheet, and Active Cell value (TM) and one above (RM)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cellTM = sheet.getActiveCell(); //sheet.getActiveCell(); \\ Ain't broke, don't fix...
  var cellRM = sheet.getActiveCell().offset(-1, 0); //sheet.getActiveCell().offset(-1, 0);
  
  // Get the value of the RM to hold
  var valueRM = cellRM.getValue();
  
  // Get the location of RM
  var locationRM = cellRM.getA1Notation();
  
  // Get the value of the TM?
  // var valueTM = cellTM.getValue();
  
  // Get A1 Notation of the TM cell
  var locationTM = cellTM.getA1Notation();  
  
  // Get the formula of the TM!
  var formulaTM = cellTM.getFormula();
  // Logger.log(formulaTM);
  
  // Set out the "if" for toAdd
  var toAdd = -5;
  
  switch(qualityOfCompletion) {
    case ("best"):
      toAdd = 10;
      break;
    case ("normal"):
      toAdd = 5;
      break;
    case ("bad"):
      toAdd = 0;
      break;
    case ("failed"):
      toAdd = -5;
      break;
      
  }
  
  
  // Fill TM Formula using locationRM
  var formulaTM = "=CEILING(" + locationRM + "*0.9,AQ5)";
  
  // Put the formula into the cell, and Add/Subtract
  cellTM.setFormula(formulaTM + "+" + toAdd);
  
}

//________________________________________

function bestTM() {
  modifyCell("best");
}


//________________________________________

function normalTM() {
  modifyCell("normal");
}

//_________________________________________

function badTM() {
  modifyCell("bad");
}

//

function failedTM() {
  modifyCell("failed");
}

//