/**
 * 
 * Googleドキュメント全体を指定したテキストで書き換える
 * FIXME:フォントの色や見出しなどの装飾は消えてしまうのに注意
 * 
 * @param  {Array.<Array.<string>>} lists - [[/置換したいテキスト/g, '置換後のテキスト']] gフラグを忘れずに
 * @return {string} 
 * 
 */
function modifyStrings(lists) {

  console.info('modifyStrings()を実行中');
  console.info('06_documentに記載');

  const document = DocumentApp.getActiveDocument();
  const body     = document.getBody();
  const original = body.getText();

  let alertString = '';
  lists.map(list => alertString += `${list[0]}　を　${list[1]}に変換する\n`);

  console.log(`置換前のテキスト`);
  console.log(original);

  const ui       = DocumentApp.getUi();
  const response = ui.alert(`置換対象を全て変換します。よろしいでしょうか？\n\n
    ${alertString}`, ui.ButtonSet.YES_NO
  );

  switch (response){
    case　ui.Button.YES:
      const replaced = lists.reduce((accumulator, list) => accumulator.replace(...list), original);

      console.log('“はい”　のボタンが押されました。');
      console.log(`変換前のテキスト`);
      console.log(replaced);

      // 文章を書き換える
      body.setText(replaced);
      break;

    case ui.Button.NO:
      console.log('“いいえ”　のボタンが押されました。');
      ui.alert('処理が中断されました。');
      break;

    default:
      console.log('閉じるボタンが押されました');
      console.log('処理が中断されました。');
      return
  }
}



/**
 * Google Documentsの内容をテキストで取得する
 * 
 * @param  {string} url - Google DocumentsのURL
 * @return {string} 
 * 
 */
function getDocContents(url) {
  const document   = DocumentApp.openByUrl(url);
  const body       = document.getBody();
  const contents   = body.getText();

  console.info('getDocContents()を実行中');
  console.info('06_documentに記載');
  console.log(contents);

  return contents
}



/**
 * Google Documentsのparagraphsを取得する
 * 
 * @param  {string} url - Google DocumentsのURL
 * @return {string} 
 * 
 */
function getDocParagraphs(url) {
  const document   = DocumentApp.openByUrl(url);
  const body       = document.getBody();
  const contents   = body.getText();
  const paragraphs = body.getParagraphs();

  console.info('getDocParagraphs()を実行中');
  console.info('06_documentに記載');

  console.log(contents);
  console.log(paragraphs);

  return paragraphs
}


/**
 * Google Documentsのparagraphsの中から正規表現に一致する箇所を抽出し、2次元配列で返す
 * ライブラリ内のgetDocParagraphsを使用すると楽ちん
 * 
 * @param  {Object.<string>} paragraphs - パラグラフ　
 * @param  {string} pattern - 正規表現の文字列
 * @return {Array.<Array.<number|string>>} 
 * 
 */
function getMatchedTextsInParagraphs(paragraphs, pattern){

  console.info('getMatchedTextsInParagraphs()を実行中');
  console.info('06_documentに記載');

  let count = 0;
  const reg = new RegExp(pattern);

  // ヘッダー行
  const values = [['該当回数', 'テキスト']];

  paragraphs.map(paragraph => {
    const array    = paragraph.getText().split('。');
    const filtered = array.filter(text => text.match(reg) !== null);
    const sentence = filtered.join(',').replace(/\r/, '');

    if(sentence.match(reg) !== null){
      count += 1;
      console.log(sentence);
      values.push([count, sentence]);
    }
  });

  console.log(values);
  console.log(`該当件数：　${count}件`);

  return values
}


/**
 * ドキュメントのPDFを作成するスクリプト
 * 
 * @param  {string} folderUrl - Google DriveのフォルダのURL
 * 
 */
function convertDocToPdf(folderUrl) {

  console.info('convertDocToPdf()を実行中');
  console.info('06_documentに記載');

  let folderId;

  // 引数のfolderUrlが定義されていない場合、入力画面を表示させてプロンプト内容を受け取る
  !folderUrl ? folderId = getFolderId(showPrompt('フォルダのURLを入力してください', 'https://drive.google.com/drive/folders/*****')) : folderId = getFolderId(folderUrl)
 
  const folder = DriveApp.getFolderById(folderId);
  console.log(`folderName: ${folder.getName()}`);

  const document = DocumentApp.getActiveDocument();
  const fileName = document.getName();
  const baseUrl  = `https://docs.google.com/document/d/${document.getId()}/export?`;
  const params   = {
    exportFormat: 'pdf',     // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       'pdf',     // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         'A4',      // 用紙サイズの指定 legal / letter / A4
    portrait:     'true',    // true → 縦向き、false → 横向き
    fitw:         'true',    // 幅を用紙に合わせるか
    docNames:     'false',   // シート名を PDF 上部に表示するか
    pagenumbers:  'true',    // ページ番号の有無
  };

  const options = Object.entries(params)
    .map(([key, value]) => `${key}=${value}`)
    .join('&');

  console.log(options);

  const token      = ScriptApp.getOAuthToken();
  const requestUrl = `${baseUrl}${options}`;

  console.log(`requestUrl:${requestUrl}`);

  const response = UrlFetchApp.fetch(requestUrl, {
    headers: {
    'Authorization': 'Bearer ' +  token
    }
  });

  const blob = response.getBlob().setName(fileName + '.pdf');
  console.log(blob);
  folder.createFile(blob);

  DocumentApp.getUi().alert(`PDF書類が下記フォルダに作成されました。\nフォルダ名：${folder.getName()}\n`);

}