/*
Get sample file ID and folder ID where you want to save your generated files.
*/
const sampleFileKey = <<<YOUR_SAMPLE_FILE_KEY_OR_ID>>>;
const folderToSaveKey = <<<YOUR_FOLDER_ID>>>;

/*
-------
Main Code Goes Below
formDoc() method is the main method here.
*/


function formDoc() {
  
  //Gets reference to sample file and target folder
  var sampleFile = DriveApp.getFileById(sampleFileKey);
  var folderToSave = DriveApp.getFolderById(folderToSaveKey);
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataKeyValue = {}
  var range = sheet.getDataRange();
  
 //gets the last filled row
  var formData = range.getValues()[sheet.getLastRow() - 1];
  //gets the form headers
  var formHeaders = range.getValues()[0];
  
  //iterates through the formdata length and assigns value based on form header
  for (var dataCounter = 0; dataCounter < formData.length; dataCounter++) {
    if (formHeaders[dataCounter].toLowerCase().includes("date")) {
      formData[dataCounter] = dtFmt(formData[dataCounter]);
    }
    dataKeyValue[formHeaders[dataCounter]] = formData[dataCounter];
  }
  
  //Receipt number is auto generated. here follows the format of ROWNUM_NAME_RECEIPTDATE
  dataKeyValue['receiptNum'] = `${sheet.getLastRow() - 1}_${dataKeyValue["Name"]}_${dataKeyValue["Receipt Date"].split('-').join('')}`;
  
  //Makes copy of the sample file and iterates throught the key value and replaces with the values captured.
  
  var document_copy = sampleFile.makeCopy(dataKeyValue['receiptNum'], folderToSave);
  var document = DocumentApp.openById(document_copy.getId());
  var document_header = document.getHeader();
  var document_body = document.getBody();

  for (var key in dataKeyValue) {
    document_header.replaceText(`<<${key}>>`, dataKeyValue[key]);
    document_body.replaceText(`<<${key}>>`, dataKeyValue[key]);
  }
  document.saveAndClose();
}

//Function to retrive desired date format
function dtFmt(datePassed) {
  var millis = Date.parse(datePassed);
  var date = new Date(millis);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  return `${day}-${month}-${year}`;
}
