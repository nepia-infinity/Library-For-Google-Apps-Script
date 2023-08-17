
/**
 * Google Slidesのテキストを取得　2次元配列で取得
 * 
 * @param  {string} url -　Google SlidesのURL
 * @return {Array.<Array.<string>>}
 * 
 */
function getSlidesContents(url){
  const presentation = SlidesApp.openByUrl(url);  
  const slides       = presentation.getSlides();

  console.log(`スライドの名前：　${presentation.getName()}`);
  console.log(`スライドの長さ：　${slides.length}`);

  let page      = 1;
  let newValues = [['ページ数', 'オブジェクトID', 'スライド内容']];

  for(const slide of slides){
    const shapes = slide.getShapes();
    for(const shape of shapes){
      const info = {
        text :      shape.getText().asString(),
        objectId:   shape.getObjectId()
      }
    console.log(`page: ${page}, ${info.text}`);
    newValues.push([page, info.objectId, info.text]);

    }
    page += 1;
  }
  console.log(newValues);
  return newValues
}




/**
 * Google Slidesのスピーカーノートを2次元配列で取得する
 * 
 * @param  {string} url - Google SlidesのURL
 * @return {Array.Array.<string>} 
 * 
 * 
 */
function getSpeakerNotes(url){
  const presentation = SlidesApp.openByUrl(url);  
  const slides = presentation.getSlides();
  let array    = [];
  
  slides.map((slide, index) => {
    const count = index + 1;
    array.push([count, slide.getNotesPage().getSpeakerNotesShape().getText().asString()]);
  });

  console.log(array);
}




/**
 * Google Slidesの全ページをjpgに変換する
 * ファイル名は、プレゼンテーション名_01.jpgのように連番で出力
 * 
 * @param {string} folderUrl - Google DriveのフォルダURL
 * 
 */
function convertSlidesToJpg(folderUrl) {
  const presentation = SlidesApp.getActivePresentation();
  const slides       = presentation.getSlides();
  const ui           = SlidesApp.getUi();

  console.log(`presentationName: ${presentation.getName()}`);
  console.log(`presentationId: ${presentation.getId()}`);
  console.log(`slides.length: ${slides.length}`);

  // UI経由でGoogle DriveのフォルダURLを入力させる
  let folderId;

  if(!folderUrl){
    const sample = 'https://drive.google.com/drive/folders/******';
    const prompt = ui.prompt('Google DriveのURLを入力してください', sample, ui.ButtonSet.OK).getResponseText();
    console.log(`input: ${prompt}`);

    folderId = getFolderId(prompt); 

  }else{
    folderId = getFolderId(folderUrl); 

  }

  const folder = DriveApp.getFolderById(folderId);
  console.log(`folderName: ${folder.getName()}`);
  
  let count = 0;

  slides.map(slide => {
    count += 1;
    const slideNumber = ('0' + count).slice(-2);

    createJpgFile_(folder, presentation, slide.getObjectId(), slideNumber)
  })
  ui.alert(`${count}件のスライドをjpgに変換しました`);
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
function createJpgFile_(folder, presentation, pageId, slideNumber){
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
