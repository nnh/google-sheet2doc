/**
* プロパティを設定する
* @param none
* @return none
*/
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('inputSheetID', ''); // スプレッドシートのID
  PropertiesService.getScriptProperties().setProperty('inputSheetSheetName', 'フォームの回答 1');
}
/**
* 列名から配列のインデックスを返す
* @param {string} column_name 列名（'A'など）
* @return Aなら0のような、getValues時の該当する配列インデックス
*/
function getColumnNumber(targetSheet, columnName){ 
  var colNumber = targetSheet.getRange(columnName + '1').getColumn();
  colNumber--;
  return colNumber;
}
function executeSs2Doc(){
  // 対象外の列を指定
  // ex.A列とC列以外の列を対象とする場合は['A', 'C']と記載する
  const exclusionColumn = ['A', 'P'];
  const targetSheetID = PropertiesService.getScriptProperties().getProperty('inputSheetID');
  const targetSheetSheetName = PropertiesService.getScriptProperties().getProperty('inputSheetSheetName');
  const targetSpreadSheet = SpreadsheetApp.openById(targetSheetID); 
  const targetSheet = targetSpreadSheet.getSheetByName(targetSheetSheetName);
  const targetSpreadSheetName = targetSpreadSheet.getName();  // SpreadSheet Name
  const outputDocsID = createOutputDocs(targetSpreadSheetName);
  const outputDocs = DocumentApp.openById(outputDocsID);
  var exclusionColumnNumber = exclusionColumn.map(col => getColumnNumber(targetSheet, col));
  // 対象外列を降順でソート
  exclusionColumnNumber.sort(function(a, b){
    if(a > b){
      return -1;
    }
    if(a < b){
      return 1;
    }
    return 0;
  });
  // スプレッドシートのデータを取得
  const lastRow = targetSheet.getLastRow();
  const lastColumn = targetSheet.getLastColumn();
  const sheetAllValues = targetSheet.getRange(1, 1, lastRow, lastColumn).getValues();
  // [重複回答]!="Y"のみ出力
  const duplicateAnswerIndex = sheetAllValues[0].indexOf('重複回答');
  const deletedDuplicateAnswerColumnValues = sheetAllValues.filter(rowValues => rowValues[duplicateAnswerIndex] != 'Y');
  // 対象外列の情報を削除
  const deletedExclusionColumnValues = deletedDuplicateAnswerColumnValues.map(function(rowValues){
    for (var i = 0; i < exclusionColumnNumber.length; i++){
      rowValues.splice(exclusionColumnNumber[i], 1);
    }
    return rowValues;
  });
  // GoogleDocsに出力する情報
  var outputValues = deletedExclusionColumnValues;
  // 見出し情報
  const headerValues = outputValues[0];
  const headerLength = headerValues.length;
  // 見出し以外の情報
  const bodyValues = outputValues.filter(function(rowValues, index){
    return index != 0;
  });
  const bodyLength = bodyValues.length;
  Logger.log(bodyValues[bodyLength - 2][headerLength - 1]);
  const outputDocsBody = outputDocs.getBody();
  // スタイルの再定義
  var stylePage = {};
  stylePage[DocumentApp.Attribute.MARGIN_LEFT] = 56.69291338582678;
  stylePage[DocumentApp.Attribute.MARGIN_RIGHT] = 56.69291338582678;
  stylePage[DocumentApp.Attribute.MARGIN_TOP] = 42.51968503937008;
  stylePage[DocumentApp.Attribute.MARGIN_BOTTOM] = 42.51968503937008;
  outputDocsBody.setAttributes(stylePage);
  var styleHeading1 = {};
  styleHeading1[DocumentApp.Attribute.FONT_SIZE] = 10;
  styleHeading1[DocumentApp.Attribute.FONT_FAMILY] = 'MS Pゴシック';
  styleHeading1[DocumentApp.Attribute.BOLD] = true;
  styleHeading1[DocumentApp.Attribute.SPACING_BEFORE] = 8;
  outputDocsBody.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, styleHeading1);
  var styleNormal = {};
  styleNormal[DocumentApp.Attribute.FONT_SIZE] = 11;
  styleNormal[DocumentApp.Attribute.FONT_FAMILY] = 'MS P明朝';
  styleNormal[DocumentApp.Attribute.BOLD] = false;
  styleNormal[DocumentApp.Attribute.LINE_SPACING] = 1.25;
  outputDocsBody.setHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL, styleNormal);
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

function createOutputDocs(targetSpreadSheetName){
  // 作成ファイル名は元のSpreadsheetのファイル名+半角スペース+西暦年月日8桁
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const createOutputDocs = DocumentApp.create(targetSpreadSheetName + ' ' + today);
  return createOutputDocs.getId();
}
