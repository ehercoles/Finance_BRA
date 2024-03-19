function confirm(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     message,
     'Are you sure you wish to proceed?',
      ui.ButtonSet.YES_NO);
  return result == ui.Button.YES;
}

// archive: not used, too slow
function deleteEmptyRows(sheet){
  var maxRows = sheet.getMaxRows(); 
  var lastRow = sheet.getLastRow();
  sheet.deleteRows(lastRow+1, maxRows-lastRow);
}

// TODO: create a new project 'Util', add as library using its script id
function getScriptID() {
  var scriptID = ScriptApp.getScriptId();

  return scriptID;
}

function logError(message) {
  MailApp.sendEmail('ehercoles@gmail.com', 'GAS error', message);
}

//a.sort(sortFunction);
function sortFunction(a, b) {
  var col = 1;
  
  if (a[col] === b[col]) {
    return 0;
  }
  else {
    return (a[col] < b[col]) ? -1 : 1;
  }
}
