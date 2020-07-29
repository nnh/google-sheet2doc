# google-sheet2doc
## 概要
GoogleSpreadsheetの内容をGoogleDocsに出力します。
## 実行前作業
- テンプレートファイルの設定  
出力したいフォルダにGoogleDocsのテンプレートファイルを保存してください。
- プロジェクトプロパティの設定  
下記箇所の('inputSheetID', '')と('templateDocumentID', '')の''にファイルのIDを入れ、functionを実行してください。 
```
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('inputSheetID', ''); // Input Spreadsheet ID
  PropertiesService.getScriptProperties().setProperty('templateDocumentID', ''); // Template Document ID
  PropertiesService.getScriptProperties().setProperty('inputSheetSheetName', 'フォームの回答 1');
}
```
## 実行方法
function executeSs2Doc()を実行してください。
## 実行結果
テンプレートファイルを保存したフォルダにGoogleDocsが出力されます。  
ファイル名は「元のSpreadsheetのファイル名+半角スペース+西暦年月日8桁」となります。  
同日に複数回実施した場合、名前が同じファイルが複数作成されます。  
