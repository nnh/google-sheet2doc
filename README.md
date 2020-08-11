# google-sheet2doc
## 概要
GoogleSpreadsheetの内容をGoogleDocsに出力します。
## 実行前作業
- テンプレートファイルの設定  
出力したいフォルダにGoogleDocsのテンプレートファイルを保存してください。
- プロジェクトプロパティの設定  
下記箇所の('inputSheetID', '')と('templateDocumentID', '')の''にファイルのIDを入れ、function registerScriptProperty()を実行してください。 
```
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('inputSheetID', '1o...J8'); // Input Spreadsheet ID
  PropertiesService.getScriptProperties().setProperty('templateDocumentID', '1E...O8'); // Template Document ID
  PropertiesService.getScriptProperties().setProperty('inputSheetSheetName', 'フォームの回答 1');
}
```
ファイルのIDはURLの'd/'と'edit...'の間の値です。  
('inputSheetID', '')には入力スプレッドシートの、('templateDocumentID', '')にはテンプレートファイルのIDを入れてください。
![input file](https://user-images.githubusercontent.com/24307469/89845234-e0c9ac00-dbb8-11ea-9d2c-df07ae48ff79.png)

## 実行方法
function executeSs2Doc()を実行してください。
## 実行結果
テンプレートファイルを保存したフォルダにGoogleDocsが出力されます。  
ファイル名は「元のSpreadsheetのファイル名+半角スペース+西暦年月日8桁」となります。  
同日に複数回実施した場合、名前が同じファイルが複数作成されます。  
