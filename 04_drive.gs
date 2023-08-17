/**
 * Googleドライブのフォルダから指定のファイル名を含むファイルを2次元配列で取得
 * 'ファイル名', 'ファイルID', 'URL', '最終更新日'を取得できる
 * 
 * @param  {string} url - GoogleドライブのフォルダのURL
 * @param  {string} query - 検索したいファイル名、省略可
 * @return {Array.<Array.<string>>}
 */
function getFilesValues(url, query){

  console.info('getFilesValues()を実行中');
  console.info('04_driveに記載');

  const files   = getDriveFiles(url);
  let newValues = [['ファイル名', 'ファイルID', 'URL', '最終更新日']];

  while (files.hasNext()) {
    const file     = files.next();
    const fileName = file.getName();

    // 2番目の引数が定義されており、ファイル名に検索queryがない場合は処理をスキップ　
    if(query && !fileName.includes(query)) continue;

    const info = {
      fileName:    fileName,
      fileId:      file.getId(),
      lastUpdated: formatDate(file.getLastUpdated(), 'yyyy/MM/dd')
    }

    // URLを追加する
    info['fileUrl'] = `https://drive.google.com/file/d/${info.fileId}/view`;

    // 引数に検索queryの指定がある場合
    if(query && fileName.includes(query)){
      console.log(`検索query：${query} に該当しました`);
      console.log(info);
      newValues.push([info.fileName, info.fileId, info.fileUrl, info.lastUpdated]);

    }else if(!query){
      console.log(info);
      newValues.push([info.fileName, info.fileId, info.fileUrl, info.lastUpdated]);

    }
  }
  console.log(newValues);
  return newValues
}



/**
 * GoogleドライブのURLからフォルダIDを抽出する
 * 
 * @param  {string} url - GoogleドライブのフォルダのURL
 * @return {string}
 */
function getFolderId(url){

  console.info('getFolderId()を実行中');
  console.info('04_driveに記載');

  let folderId;
  const reg = /.*\//;

  // URLから不要な文字列を削除する
  if(url.match(reg) !== null){
    folderId = url
    .replace(reg, '')
    .replace(/.hl=.*/, '');

    console.log(`folderId: ${folderId}`);
  }

  return folderId
}



/**
 * Googleドライブのフォルダ内のファイルのオーナー権限を一括で指定の相手に譲渡する
 * FIXME: 個人からビジネスドメインへの権限移譲は失敗することがある
 * FIXME: 個人アカウントには、ファイルのオーナー権限を譲渡という概念がないので失敗する
 * 
 * @param  {string} url - GoogleドライブのフォルダのURL
 * @param  {string} accountId - ファイルを譲渡したいアカウントID
 * @return {Array.<Array.<string>>}
 */
function transferOwnership(url, accountId){

  console.info('transferOwnership()を実行中');
  console.info('04_driveに記載');

  const files = getDriveFiles(url);

  while (files.hasNext()) {
    const file = files.next();
    console.log(`ファイル名：${file.getName()}`);

    try{
      file.setOwner(accountId).setDescription('一括で、オーナー権限を移すテスト');

    }catch{
      console.warn(`ファイル譲渡に失敗しました。`);
      console.log(`個人からビジネスドメインへの権限移譲は制限されている可能性あります。`);
      console.log(`もしくは、ファイルへのアクセスが制限されている可能性があります`);

    }
  }
}



/**
 * Googleドライブの特定フォルダ内のファイルを取得する
 * 
 * @param  {string} url - GoogleドライブのフォルダのURL
 * @return {Object.<Object.<string>>}
 */
function getDriveFiles(url){

  const folderId = getFolderId(url);

  console.info('getDriveFiles()を実行中');
  console.info('04_driveに記載');

  const folder   = DriveApp.getFolderById(folderId);
  const files    = folder.getFiles();

  console.log(`取得対象のフォルダ：　${folder.getName()}`);

  return files
}


/*
共有範囲
DriveApp.getFileById('File_ID').getSharingAccess();

ANYONE: ウェブ上で一般公開
ANYONE_WITH_LINK: リンクを知っている全員
DOMAIN: DOMAIN内の全員
DOMAIN_WITH_LINK: DOMAIN内でリンクを知っている全員
PRIVATE: 特定のユーザ

共有範囲の権限
DriveApp.getFileById('File_ID').getSharingPermission();
VIEW: 閲覧者
EDIT: 編集者
COMMENT: コメント可
OWNER: オーナー
ORGANIZER: オーガナイザー
NONE: なし
*/

/**
 * 指定したユーザーにフォルダの閲覧権限や編集権限を一括で付与する
 * FIXME: 一度編集権限を付与してしまうとアクセス権を剥奪しないと閲覧オンリーなどに切り替えることが出来ない
 * 
 * @param  {string} url - GoogleドライブのフォルダのURL
 * @param  {string} users - GoogleドライブのフォルダのURL
 * @param  {string} role - 編集権限
 * 
 */
function authorizeEditing(url, users, role){

  console.info('authorizeEditing()を実行中');
  console.info('04_driveに記載');

  const folderId = getFolderId(url);
  const folder   = DriveApp.getFolderById(folderId);
  const reg      = / gmail.* | icloud.* /;

  let permission;

  if(role === '編集'){
    //型が文字列ではない事に注意が必要
    permission = DriveApp.Permission.EDIT;

  }else{
    role = '閲覧権限';
    permission = DriveApp.Permission.VIEW;

  }

  for(const user of users){

    if(user.match(reg) !== null){
      // 個人のフリーアドレス宛に権限を付与する
      console.log(`${user} に ${folder.getName()}　の　${role} 権限を付与しました`);
      folder.addEditor(user).setSharing(DriveApp.Access.PRIVATE, permission);

    }else{
      // ドメイン内のユーザーに引数に応じた権限を付与する
      console.log(`${user} に ${folder.getName()}　の　${role} 権限を付与しました`);
      folder.addEditor(user).setSharing(DriveApp.Access.ANYONE_WITH_LINK, permission);

    }
  }
  console.log(`全員に権限の付与が完了しました`);
}


/**
 * 1次元配列の内容を元にフォルダを作成する
 * 
 * @param  {string} url - GoogleドライブのフォルダのURL
 * @param  {Array.<string>} newFolderNameList - 1次元配列　この配列に格納された値がフォルダ名となる
 * @return {Array.<string>} innerFolderNameList - 1次元配列　さらに内側にフォルダを作成する場合に指定する
 */
function createFolders(url, newFolderNameList, innerFolderNameList) {

  console.info('createFolders()を実行中');
  console.info('04_driveに記載');

  const folderId = getFolderId(url);
  const folder   = DriveApp.getFolderById(folderId);

  console.log(`対象フォルダー名：${folder.getName()}`);
  console.log(`${newFolderNameList.length}個のフォルダーを新規作成します。`);

  for(const newFolderName of newFolderNameList){
    const newFolder = folder.createFolder(newFolderName);
    console.log(`${newFolderName}というフォルダが新規作成されました`);

    if(innerFolderNameList){
      for(const innerFolderName of innerFolderNameList){
        newFolder.createFolder(innerFolderName);
        console.log(`${newFolderName}の中に、${innerFolderName}というフォルダが新規作成されました`);
      }
    }
  }
  console.log(`全てのフォルダの作成が終了しました。作成結果が反映されるまで少し時間が掛かる場合があります`);

}



/**
 * 
 * スプレッドシートのURLと印刷範囲をして、PDFを生成する
 * 
 * @param  {string} sheetUrl - シートのURL
 * @param  {string} stringRange - 'A2:D30' などPDFの生成範囲
 * @param  {string} folderUrl - Google DriveのフォルダURL
 * @return {Object.<string>}
 * 
 */
function convertSheetToPdf(sheetUrl, stringRange, folderUrl) {

  console.info('convertSheetToPdf()を実行中');
  console.info('04_driveに記載');

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByUrl(sheetUrl);
  const info  = {
    spreadsheetId: spreadsheet.getId(),
    sheetId:       sheet.getSheetId(),
    printArea:     stringRange.replace(':', '%3A')
  }

  console.log(info);

  // 印刷情報、サイズや紙の向きや範囲などの情報を付与する
  const exportOptions = {
    format:       'pdf',               // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         'A4',                // 用紙サイズの指定 legal / letter / A4
    portrait:     'true',              // true → 縦向き、false → 横向き
    fitw:         'true',              // 幅を用紙に合わせるか
    sheetnames:   'false',             // シート名を PDF 上部に表示するか
    printtitle:   'false',             // スプレッドシート名を PDF 上部に表示するか
    pagenumbers:  'false',             // ページ番号の有無
    gridlines:    'false',             // グリッドラインの表示有無
    fzr:          'false',             // 固定行の表示有無
    range:        info.printArea, // 対象範囲「%3A」 = : (コロン)  
    gid:          info.sheetId    // シート ID を指定 (省略する場合、すべてのシートをダウンロード)
  };

  console.log(exportOptions);
 
  const array = [];
  for(const [key, value] of Object.entries(exportOptions)) {
    array.push(`${key}=${value}`);
  }
  console.log(array);

  const fileUrl = 'https://docs.google.com/spreadsheets/d/'+ info.spreadsheetId +'/export?';
  const options = array.join('&');
  console.log(options);

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(fileUrl + options, {
    headers: {
    'Authorization': 'Bearer ' +  token
    }
  });

  const today    = formatDate(new Date(), 'yyyy_MMdd_HH:mm');
  const fileName = `${sheet.getName()}_${today}`;
  const blob     = response.getBlob().setName(fileName + '.pdf');

  //　フォルダを指定してファイルを保存する
  const folderId = folderUrl.replace(/.*\//, '');
  const folder   = DriveApp.getFolderById(folderId);
  folder.createFile(blob);
 
}//end



/**
 * Google Driveのフォルダ内のファイル名を任意の文字列にリネームする
 * 
 * @param {string} url - フォルダのURL
 * @param {Array.<Array.<string>>} values - [['置換対象', '置換後']]
 * 
 */
function renameAllFile(url, values){

  console.info('renameAllFile()を実行中');
  console.info('04_driveに記載');

  const folderId = getFolderId(url);
  const folder   = DriveApp.getFolderById(folderId);
  const files    = folder.getFiles();

  console.log(`取得対象フォルダ名：${folder.getName()}`);

  let count = 0;
  let lists = [
    [/\s/, ''],
    ['★', ''],
  ];

  if(values){ 
    lists = lists.concat(values);
    console.log(`結合後`);
    console.log(lists); 
  }

  let newValues = [];

  while (files.hasNext()) {
    const file     = files.next();
    const original = file.getName();

    count += 1;
    console.log(`変更前: ${count}. ${original}`);

    const newFileName = lists.reduce((accumulator, list) => accumulator.replace(...list), original);
    console.log(`変更後: ${count}. ${newFileName}`);
    
    newValues.push([newFileName, file.getId()]);

  }
  console.log(`${count}件`);
  console.log(newValues);

  showAlertBeforeExecution_(newValues);

}


/**
 * 実行前にアラートを表示する
 * 
 * @param {Array.<Array.<string>>} values - [['ファイル名', 'ファイルID']]
 * 
 */
function showAlertBeforeExecution_(values){

  console.info('showAlertBeforeExecution_()を実行中');
  console.info('04_driveに記載');

  let string = '';
  values.map((row, index) => {
    const number = index + 1;
    string += `${number}. ${row[0]}\n`;
  });

  let fileIdArray = [];
  values.map(row => fileIdArray.push(`${row[1]}`));
  console.log(fileIdArray);

  const ui       = SpreadsheetApp.getUi();
  const response = ui.alert(`
    ファイル名を下記のように変更してもよろしいですか？\n\n
    ${string}`, ui.ButtonSet.YES_NO
  );

  switch (response){
    case ui.Button.YES:
      console.log('“はい” のボタンが押されました。');
      fileIdArray.map((fileId, index) => DriveApp.getFileById(fileId).setName(values[index][0]));
      ui.alert('ファイル名の変更が完了しました。');
      break;

    case ui.Button.NO:
      console.log('“いいえ” のボタンが押されました。');
      ui.alert('処理が中断されました。');
      break;

    default:
      console.log('処理が中断されました。');
      return
  }
}



/**
 * リンク付きのファイル一覧をHTMLとして表示する
 * 
 * @param  {string} url - フォルダのURL
 * @return {string} string - ファイル一覧のHTMLタグ
 */
function getFileNameWithUrl(url){

  console.info('getFileWithUrl()を実行中');
  console.info('04_driveに記載');

  const folderId = getFolderId(url);
  const folder   = DriveApp.getFolderById(folderId);
  const files    = folder.getFiles();

  console.log(`取得対象フォルダ：${folder.getName()}`);

  let html = `<p>フォルダ名：　<a href="${url}">${folder.getName()}</a></p>`;

  while (files.hasNext()) {
    const file     = files.next();
    const fileName = file.getName();
    const fileUrl  = `https://drive.google.com/file/d/${file.getId()}/view`;
    console.log(fileUrl);
    
    html += `<p><a href="${fileUrl}">${fileName}</a></p>`;

  }
  console.log(html);
  showHtmlSentence(html, 'ファイル一覧を表示します');
}



/**
* フォルダから画像ファイルのみを取得し、ファイル名でsortし2次元配列で返す
* [['ファイル名_01.jpg', 'ファイルID'], ['ファイル名_02.jpg', 'ファイルID']]
* 22. getDriveFiles()と挙動が近いが、こちらは画像ファイルのみを配列に加える
* 
* @param  {string} folderUrl - Google DriveのフォルダのURL
* @return {Array.<Array.<string>>}
* 
*/
function getImageFiles_(folderUrl) {
  const folderId = getFolderId(folderUrl);
  const folder   = DriveApp.getFolderById(folderId);
  const files    = folder.getFiles();
  const values   = [];

  //jpg、gif、pngを取得してシートの最終行に挿入する
  while(files.hasNext()){
    const file = files.next();
    if(file.getMimeType().match(/^image\/(?:jpeg|gif|png)$/i)) {
      values.push([file.getName(), file.getId()]);
    }
  }

  //ファイル名を連番順にSORTする
  values.sort((a, b) => a[0] < b[0] ? -1:1);
  console.log(values);

  return values

}

