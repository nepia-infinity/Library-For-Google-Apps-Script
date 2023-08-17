
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
  const sheetName   = 'ScriptDetail';
  const activeSheet = SpreadsheetApp.getActiveSheet();

  if(activeSheet.getName() !== sheetName){
    console.log(`処理対象のシートではないため、処理を終了`);
    return
  }

  const values     = activeSheet.getDataRange().getValues();
  const column     = generateHeaderIndex(values);
  const activeCell = getActiveCell(activeSheet);

  if(activeCell.column !== column.scriptName + 1){
    console.log(`処理対象列ではないため、処理を終了`);
    return

  }else if(activeCell.row === 1){
    console.log(`見出し行なので処理終了します。`);
    return

  }else{
    const range = activeSheet.getRange(activeCell.row, column.updateDate + 1);
    console.log(`転記範囲:　${range.getA1Notation()}`);

    const today = formatDate(new Date(), 'yyyy/MM/dd');
    range.setValue(today);
  }
}



