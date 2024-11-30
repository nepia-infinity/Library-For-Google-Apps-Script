/**
 * Google Slidesのテキストを取得　2次元配列で取得
 * 
 * @param  {string} url -　Google SlidesのURL
 * @return {Array.<Array.<string>>}
 * 
 * newValues.unshift(['ページ数', 'オブジェクトID', 'スライド内容']);
 * 
 */
function getSlidesContents(url) {

  console.info(`getSlidesContents()を実行中`);
  console.info(`07_slides()に記載`);

  const presentation = SlidesApp.openByUrl(url);
  const slides = presentation.getSlides();

  console.log(`スライドの名前: ${presentation.getName()}`);
  console.log(`スライドの長さ: ${slides.length}`);

  const newValues = slides.map((slide, index) => {
    return slide.getShapes().map(shape => [index + 1, shape.getText().asString(), shape.getObjectId()]);
  });

  newValues.unshift(['ページ数', 'オブジェクトID', 'スライド内容']);

  console.log(newValues);
  return newValues;
}



/**
 * Google Slidesのスピーカーノートを2次元配列で取得する
 * 
 * @param  {string} url - Google SlidesのURL
 * @return {Array.Array.<string>} 
 * 
 */
function getSpeakerNotes(url){
  const presentation = SlidesApp.openByUrl(url);
  const slides = presentation.getSlides();  
  const values = slides.map((slide, index) => {
    return [index + 1, slide.getNotesPage().getSpeakerNotesShape().getText().asString()];
  });

  console.log(values);
  return values
}



/**
 * Google Slidesの全ページをjpgに変換する
 * ファイル名は、プレゼンテーション名_01.jpgのように連番で出力
 * 
 * @param {string} folderUrl - Google DriveのフォルダURL
 * 
 */
function convertSlidesToJpg(folderUrl) {

  console.info(`convertSlidesToJpg()を実行中`);
  console.info(`07_slidesに記載`);

  const presentation = SlidesApp.getActivePresentation();
  const slides       = presentation.getSlides();
  const ui           = SlidesApp.getUi();

  console.log(`presentationName: ${presentation.getName()}`);
  console.log(`presentationId:   ${presentation.getId()}`);

  const folderId = getFolderId(folderUrl);
  const folder   = DriveApp.getFolderById(folderId);
  console.log(`folderName: ${folder.getName()}`);

  slides.forEach((slide, index) => {
    const count = index + 1;
    const slideNumber = ('0' + count).slice(-2);
    createImgeFromSlide_(folder, presentation, slide.getObjectId(), slideNumber);
  });

  ui.alert(`${slides.length}件のスライドをjpgに変換しました`);
  return
}



/**
 * Google Slidesの各ページをjpgとして指定したフォルダに保存する
 * 
 * @param {Object.<Array.<string>>} folder - フォルダオブジェクト
 * @param {Object.<Array.<string>>} presentation - Google Slidesのスライドオブジェクト
 * @param {string} pageId - Google SlidesのスライドID
 * @param {string} slideNumber - ファイルの命名に使用する連番
 * 
 */
function createImgeFromSlide_(folder, presentation, pageId, slideNumber){

  console.info(`createImgeFromSlide_()を実行中`);
  console.info(`07_slidesに記載`);

  const presentationId = presentation.getId();
  const baseUrl        = `https://docs.google.com/presentation/d/`;
  const requestUrl     = `${baseUrl}${presentationId}/export/jpeg?id=${presentationId}&pageid=${pageId}`;
  console.log(`requestUrl:${requestUrl}`);

  const options = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };

  const response     = UrlFetchApp.fetch(requestUrl, options);
  const responseCode = response.getResponseCode();
  console.log(`responseCode: ${responseCode}`);

  const fileName = `${presentation.getName()}_${slideNumber}.jpg`;
  console.log(`fileName: ${fileName}`);
  response.getResponseCode() === 200 ? folder.createFile(response.getBlob()).setName(fileName) : console.warn(作成に失敗しました)

  return

}



/**
 * 
 * スライドに表を追加する
 * @param  {string} sheetUrl - スプレッドシートのURL
 * 
 * NOTICE: activate Google Slides API
 * 
 */
function insertTable(sheetUrl, ...params) {

  console.info(`insertTable()を実行中`);
  console.info(`07_slidesに記載`);

  const slides   = SlidesApp.getActivePresentation();
  const original = getValues(sheetUrl);

  // paramsが定義されている場合は、その引数に該当する値のみを表から取得する
  const values   = params ? getFilteredValues(original, params) : original
  const newSlide = slides.appendSlide();

  console.log(`追加したスライドにテーブルが挿入されます`);

  // 最初に表のみを作成して、順に文字列のみを挿入していく
  const table = newSlide.insertTable(values.length, values[0].length);
  values.forEach((record, rowIndex) => record.forEach((value, columnIndex) => {
    table.getCell(rowIndex, columnIndex).getText().setText(value.toString());
  }));
}



/**
 * 指定したスライドインデックスに存在するテーブル内の文字列を中央に配置する
 * 垂直方向の中央にすることは出来なかった
 * 
 * @param  {number} slidesIndex - スライドのページ 0始まりなので2ページ目の場合、1と指定
 * 
 * 
 */
function centerTextInCell(slidesIndex) {

  console.info(`centerTextInCell()を実行中`);
  console.info(`07_slidesに記載`);

  const slides = SlidesApp.getActivePresentation();
  const page   = slides.getSlides()[slidesIndex];
  
  const table = page.getTables()[0];
  const info  = {
    numRows:    table.getNumRows(),
    numColumns: table.getNumColumns()
  }
  console.log(info);

  for(let rowIndex = 0; rowIndex < info.numRows; rowIndex++){ // 行
    for(let columnIndex = 0; columnIndex < info.numColumns; columnIndex++){ // 列
      const currentCell = table.getCell(rowIndex, columnIndex);
      const paragraph   = currentCell.getText().getParagraphs()[0];

      const style = paragraph.getRange().getParagraphStyle();
      style.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      console.log(`rowIndex: ${rowIndex}, columnIndex: ${columnIndex}, value: ${currentCell.getText().asString()}`);
      
    }
  }
}
