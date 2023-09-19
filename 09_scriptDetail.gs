
/**
 * ライブラリ使用に必要なスクリプトIDを表示する
 * 
 * 
 */
function showScriptId(){
  const scriptId = '<p>18rg2maFYXNmPmB2R-8s3UuFG850j5OLw4WBvcOrghzRMlfVbQWDgOVvZ</p>';
  const html     = HtmlService.createHtmlOutput(scriptId)
  .setWidth(800)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'ライブラリ利用に必要なスクリプトID');
}

/**
 * ScriptDetailの表を更新すると、本日の日付を転記する
 * 
 */
function onEdit(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const sheet  = getActiveSheetByUrl(url);
  const values = sheet.getDataRange().getValues();
  const activeCell = getActiveCell(sheet);

  let column = generateHeadersIndex(values, 0);
  column     = modifyObject(column);

  const today = formatDate(new Date(), 'yyyy/MM/dd');

  //編集されている列が、column.scriptName、2列目かつ見出し行ではない場合
  activeCell.column === column.scriptName && activeCell.row !== 1 ? sheet.getRange(activeCell.row, column.updateDate).setValue(today) : false
    
}




