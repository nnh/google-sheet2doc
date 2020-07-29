/**
* Output the contents of the spreadsheet to Google Docs
* @param none
* @return none
*/
function executeSs2Doc(){
  // Exclude Columns
  // e.g. If you want to target columns other than columns 'A' and 'C', ['A', 'C']
  const exclusionColumn = ['A', 'P'];
  const targetSpreadSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('inputSheetID')); 
  const targetSheet = targetSpreadSheet.getSheetByName(PropertiesService.getScriptProperties().getProperty('inputSheetSheetName'));
  const targetSpreadSheetName = targetSpreadSheet.getName();  // Input SpreadSheet Name
  const outputDocsID = createOutputDocs(targetSpreadSheetName);
  const outputDocs = DocumentApp.openById(outputDocsID);  // Output Google Document
  // Get the array index from the column names and sort them in descending order
  // (In order to remove the exclude column from the last column)
  var exclusionColumnNumber = exclusionColumn.map(col => getColumnNumber(targetSheet, col));
  exclusionColumnNumber.sort(function(a, b){
    if(a > b) { return -1; }
    if(a < b) { return 1; }
    return 0;
  });
  // Get data from the spreadsheet
  const sheetAllValues = targetSheet.getRange(1, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).getValues();
  // Output '[重複回答]!="Y"' only
  const duplicateAnswerIndex = sheetAllValues[0].indexOf('重複回答');
  const deletedDuplicateAnswerColumnValues = sheetAllValues.filter(rowValues => rowValues[duplicateAnswerIndex] != 'Y');
  // Remove the exclude columns
  const deletedExclusionColumnValues = deletedDuplicateAnswerColumnValues.map(function(rowValues){
    for (var i = 0; i < exclusionColumnNumber.length; i++){
      rowValues.splice(exclusionColumnNumber[i], 1);
    }
    return rowValues;
  });
  // Values to output to GoogleDocs
  var outputValues = deletedExclusionColumnValues.map(function(x){
    var temp = x.map(val => String(val));  // Convert all of them to strings
    return temp;
  });
  // Header
  const headerValues = outputValues[0];
  const headerLength = headerValues.length;
  // Answer
  const bodyValues = outputValues.filter(function(rowValues, index){
    return index != 0;
  });
  const bodyLength = bodyValues.length;
  const outputDocsBody = outputDocs.getBody();
  var outputRow = -1;
  for (var i = 0; i < bodyLength; i++){
    for (var j = 0; j < headerLength; j++){
      var headerParagraph;
      var bodyParagraph;
      outputRow++;
      headerParagraph = outputDocsBody.insertParagraph(outputRow, headerValues[j]);
      outputRow++;
      bodyParagraph = outputDocsBody.insertParagraph(outputRow, bodyValues[i][j]);
      headerParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      bodyParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }
    if (i != (bodyLength - 1)){
      outputRow++;
      outputDocsBody.insertPageBreak(outputRow);
    }
  }  
}
/**
* Creating a Google Document to output
* @param {string} The name of the original Spreadsheet file
* @return {string} The ID of the GoogleDocument that was created
*/
function createOutputDocs(targetSpreadSheetName){
  // Created file name is the name of the original Spreadsheet file + one-byte space + 8-digit date
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  // Copy the template file
  const templateFile = DriveApp.getFileById(PropertiesService.getScriptProperties().getProperty('templateDocumentID'));
  const createOutputDocs = templateFile.makeCopy(targetSpreadSheetName + ' ' + today);
  return createOutputDocs.getId();
}
/**
* Setting Project Properties
* @param none
* @return none
*/
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('inputSheetID', ''); // Input Spreadsheet ID
  PropertiesService.getScriptProperties().setProperty('templateDocumentID', ''); // Template Document ID
  PropertiesService.getScriptProperties().setProperty('inputSheetSheetName', 'フォームの回答 1');
}
/**
* Returns an array index from a column name
* @param {string} column_name (e.g. 'A')
* @return {number} The corresponding array index of getValues (e.g. return 0 if 'A')
*/
function getColumnNumber(targetSheet, columnName){ 
  var colNumber = targetSheet.getRange(columnName + '1').getColumn();
  colNumber--;
  return colNumber;
}
