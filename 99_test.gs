/**
 * スプレッドシートのURLからsheetオブジェクトを取得する
 * 
 */
function test_getSheetByUrl() {
  const url = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  getSheetByUrl(url);
  // getSheetByUrl(url, 'sheetName');
}

/**
 * Rangeオブジェクトを取得する
 * 
 */
function test_getRange(){
  const url   = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const sheet = getActiveSheetByUrl(url);
  const info  = {row: 2, column: 1}
  getRange(sheet, info);
  getRange(sheet, 'A2:G69');
}


/**
 * n列目の文字が入力されている最終行を取得する
 * 
 */
function text_getLastRowWithText(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getValues();
  getLastRowWithText(values, 2);
}


/**
 * Dateオブジェクトを指定した表示形式に変換する
 * 
 * 
 */
function test_formatDate(){
  const date = new Date();
  formatDate(date, 'yyyy/MM/dd (E)');
  //formatDate_(date, 'yyyy/MM/dd');

}


/**
 * 
 * 文字列から正規表現を指定して、置換後の文字列を取得する
 * 
 */
function test_extractText(){
  const string = 'ID:564321 波風ミナト';
  extractText(string, /ID:[0-9]{6}/, [/ID.*:/, /\s.*/]); //ID:564321
}



/**
 * 
 * アクティブなセルを取得する
 * 
 */
function test_getActiveCell(){
  const sheet = SpreadsheetApp.getActiveSheet();
  getActiveCell(sheet);
}


/**
 * 
 * ヘッダー行をオブジェクトで取得する
 * 
 */
function test_replaceHeadersValues(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=618215393';
  const sheet  = SpreadsheetApp.openByUrl(url).getSheetByName('test');
  const values = sheet.getDataRange().getValues();
  const column = {id: 'ID', name: '名前', branch: '部署', address: '住所'};
  replaceHeadersValues(values, 0, column);

}



/**
 * 
 * ヘッダー行をオブジェクトで取得する
 * 
 */
function test_generateHeadersIndex(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const sheet  = SpreadsheetApp.openByUrl(url).getSheetByName('ScriptDetail');
  const values = sheet.getDataRange().getValues();
  generateHeadersIndex(values, 0);

}




function test_showPrompt(){
  showPrompt('何かを入力してください', 'あああああ');
}


/**
 * 
 * URLを取得して、シートの情報を2次元配列で取得する
 * 
 */
function test_getValues(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const values = getValues(url);

  //引数が1つの場合
  getFilteredValues(values, 'showPrompt()');

  //引数が複数かつ全てが含まれる場合
  getFilteredValues(values, '2次元配列', '01_spreadsheet');
}



function test_getFilteredValues(){
  const array    = ['28', 'setValues()', '2次元配列を転記する', '01_spreadsheet', 'test_setValues()', '2023/06/18', ''];
  const params   = ['2次元配列', '01_spreadsheet'];

  const text  = '2次元配列を転記する';
  console.log(text);
  console.log('text.includes(params[0])',text.includes(params[0])); //true
  console.log('array.includes(params[0])', array.includes(params[0]), '部分一致の場合false');

  const text2 = '01_spreadsheet';
  console.log(text2);
  console.log('text2.includes(params[1])', text2.includes(params[1])); //true
  console.log('array.includes(params[1])', array.includes(params[1])); //完全一致の場合はtrue

  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getValues();

  getFilteredValues(values, '2次元配列', '01_spreadsheet');
}

/**
 * 
 * オブジェクトが格納された配列から、特定のkeyのみを残して新しい配列を作成する
 * 
 */
function test_reduceObjectKeys() {

  const values = [
    {id: 0, name: 'nobita', address: 'chiba', mail: 'sample@docom.jp', hobby: 'guitar'},
    {id: 1, name: 'shizuka', address: 'tokyo', mail: 'sample@softbank.jp', hobby: 'piano'},
  ];

  reduceObjectKeys(values, 'name', 'hobby');
  
}


/**
 * 
 * spliceメソッド
 * 
 */
function test_splice(){
  const url    = 'https://docs.google.com/spreadsheets/d/1zY2Dt_xmsYwJEAGwhVa7nK4_C29r1p1mag8DPweb8Hw/edit#gid=0';
  const values = getValues(url);
  const array  = generateArray(values, 1);

  const lastIndex = array.length;
  splice(array, lastIndex, 0, 'うどん県');

}


/**
 * 
 * 引数に指定した日付までの日付を生成する
 * 
 */
function test_dateStringValues(){
  const targetDate = '2024/03/02';
  generateDateStringValues(targetDate);
}


/**
 * 
 * 配列から重複した値を削除する
 * 
 */
function test_removeDuplicates(){
  const array = ['aaa', 'bbb', 'ddd', 'ccc', 'aaa', 'ddd', 'eee', 'aaa'];
  removeDuplicates(array);
}


/**
 * 
 * スプレッドシートもしくはシートから検索キーワードを含むセルを返す
 * 
 */
function test_createTextFinder(){
  const url = 'https://docs.google.com/spreadsheets/d/1v6AqBZ-Y7wVrFjVAhzeEuROw1Qb0VhiFNzrHC9xiC0w/edit#gid=1125887426';
  createTextFinder(url, 'PDF');
  createTextFinder(url, 'PDF', 'RAW');
}


/**
 * 
 * 指定したカレンダーの予定を全て取得する
 * 
 */
function test_getCalAllEvents(){
  getCalAllEvents('2023/09/10', 1);
}


/**
 * 
 * フォルダを指定して、ファイル名などを2次元配列として取得する
 * 
 */
function test_getDriveFiles(){
  const url = 'https://drive.google.com/drive/folders/1hTaUoYGwl8mnFIUF0dd6GkhVucb18dmY';
  getFilesValues(url, 'ミラ2チェア');
  
}



/**
 * 
 * 2次元配列の中から不要な列を取り除く 
 * 
 */
function test_selectColumns(){
  const url    = 'https://drive.google.com/drive/folders/1hTaUoYGwl8mnFIUF0dd6GkhVucb18dmY';
  const values = getFilesValues(url, 'ミラ2チェア');
  const column = {fileName: 0, fileURL: 2};
  selectColumns(values, column);
  
}



/**
 * 
 * Gmailからスレッドを取得する
 * 
 */
function test_getGmailThread(){
  const values = getGmailThread('お使いの Google アカウントへのアクセス', 50);
  const column = {date: 0, subject: 2};
  selectColumns(values, column)
}


/**
 * 
 * オブジェクトの中身全てに1を足す
 * 
 */
function test_modifyObject(){
  const original = {date: 0, subject: 2};
  modifyObject(original);
}


/**
 * 
 * オブジェクトの中身を上書きする
 * すごく似た関数がある
 */
function test_sortInsideObject(){
  const original = {subject: 2, date: 4, id: 1};
  sortInsideObject(original, 1);
  swapWithAscendingIndex(original);
}


/** sortの挙動補足 */
function test_sortInsideObject_sideNote(){
  const values = [['id', 5],['name', 3],['department', 4]];
  const index  = 1;
  values.sort((previous, current) => (previous[index] < current[index]) ? -1 : 1);

  // 1週目：previous = undefined,   current = ['id', 5] 
  // 2週目：previous = ['id', 5],   current = ['name', 3]
  // 3週目：previous = ['name', 3], current = ['department', 4]]
}


/**
 * 
 * 値が配列の中に存在するかどうかを調べる
 * 
 */
function test_isNewValue(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=476791012';
  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getValues(); // 転記済みの情報　2次元配列
  console.log(values);

  const idsArray  = generateArray(values, 0);
  const newValues = [
    [ 'ID', '応募者', '受信日' ],
    [ 'JP11225', '牧瀬紅莉栖', '2023/04/23'],
    [ 'JP18611', '岡部倫太郎', '2023/04/23' ],
    [ 'JP27659', '橋田至', '2023/04/30' ]
  ];

  isNewValue(idsArray, newValues, 0);
  
}


/**
 * 
 * フォルダを大量に作成する
 * 
 */
function test_createFolders(){
  const url = showPrompt('フォルダのURLを入力してください', 'https://drive.google.com/drive/folders/*****');

  const folderNameList      = ['folderA', 'folderB', 'folderC'];
  const innerFolderNameList = ['1', '2', '3'];

  createFolders(url, folderNameList, innerFolderNameList);
}


/**
 * 
 * setValuesのテスト　2次元配列を転記する
 * 
 */
function test_setValues(){
  const url      = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=476791012';
  const values   = [['JP55475', '炭治郎', '東京']];
  const sheet    = getSheetByUrl(url);
  const original = sheet.getDataRange().getValues();

  const targetRow = getLastRowWithText(original, 1) + 1;
  const info      = {row: targetRow, column: 1}
  setValues(sheet, info, values, true);
  // setValues(sheet, info, values);
}


/**
 * Meet URL付きで予定を登録する
 * createEventsFromSheetValuesのテスト
 * 
 * 
 */
function test_createEvents(){
  console.log(CalendarApp.getDefaultCalendar());
  console.log(CalendarApp.getDefaultCalendar().getId());
  const url = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=1416056032';
  registerEvents(url, 0);
}


/**
 * カレンダーから指定した予定を削除する
 * 
 */
function test_deleteEvents(){
  deleteEvents();
}


/**
 * 1から100までの整数を生成し、1次元配列で返す
 * 
 */
function test_generateNumbers(){
  const start = 1;
  const end   = 100;

  console.log(`${start}から${end}までの整数を生成する`);
  const array = generateNumbers(start, end);

  const random = getRandomNumber(end);
  console.log(`ランダムな整数：　${random}`);

  findLargestNumber(array);

}


/**苗字を返すスクリプトのテスト*/
function test_getLastName(){
  getLastName('野比　のび太');
  getLastName('坂田銀時'); //OK
  getLastName('西園寺姫奈'); //OK
  getLastName('長久木勇二'); //OK
  getLastName('長谷川泰三（マダオ）'); //OK
  getLastName('嵯峨野芳輝'); //OK
}


/**携帯電話を返す関数 */
function test_getCellPhoneNumber(){
  const string  = '5012345678';
  const string2 = '8034567793';
  const string3 = '09012568125';
  const string4 = 8089565025;

  convertCellPhoneNumber(string);
  convertCellPhoneNumber(string2);
  convertCellPhoneNumber(string3);
  convertCellPhoneNumber(string4);
}


/**スライドの内容を取得 */
function test_getSlidesContents(){
  const url = 'https://docs.google.com/presentation/d/11uXs1bln84lhx9lPMPEeQBQw4ei5O6zdR-dYUtEuyCI/edit#slide=id.gc6f8954bc_0_53';
  getSlidesContents(url);
}

/**イベントIDから予定の詳細を取得する */
function test_getInfoByEventId(){
  const eventId = 'ektieo6tv8g0cmrjauea4ni3k8';
  getInfoByEventId(eventId);
}

/**月次のデータを作成する */
function test_createMonthlyDataNames(){
  createMonthlyDataNames('2023年', 12, '月');
}

/**既存の配列に他の配列の要素を追加する */
function test_UnshiftGuests() {
  const guestEmails = ['guest1@example.com', 'guest2@example.com'];
  const organizers = ['organizer1@example.com', 'organizer2@example.com'];

  console.log('Before unshift:');
  console.log('Guest Emails:', guestEmails);

  guestEmails.unshift(...organizers);

  console.log('\nAfter unshift:');
  console.log('Guest Emails:', guestEmails);
}


/**予定を編集する */
function test_editEvents(){
  editEvents('日時を編集する', 0);
  // editEvents('詳細欄を編集する', 0);
}


function test_showEditEventsLauncher(){
  showEditEventsLauncher();
}


/**予定の作成者を取得する */
function test_getCreator(){
  // https://developers.google.com/apps-script/reference/calendar/calendar-event-series?hl=ja#getcreators
  const event = CalendarApp.getEventById('ektieo6tv8g0cmrjauea4ni3k8');
  console.log(event.getCreators());
  console.log(event.getCreators()[0], typeof event.getCreators()[0]);
}

/**日付の配列を作成する */
function test_generateStringDate(){
  generateDateStringValues('2023/08/10');
}

/**編集権限を付与する */
function test_grantEditPermissionToFolder(){
  grantEditPermissionToFolder('https://drive.google.com/drive/folders/1rhbnV8lGbrI6QvOw-7JEPZY5UkgZaMNM', ['nepia_infinity@icloud.com'], '編集');
}


/**1から1000までの整数を生成 */
function test_generateNumbersArray(){
  generateNumbers(1, 1000);
}


/**シートからPDFを作成 */
function test_convertSheetToPdf(){
  const sheetUrl  = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const folderUrl = 'https://drive.google.com/drive/folders/1So21XbYuC5YHmSVmilOcfManpGRALMpO';
  convertSheetToPdf(sheetUrl, 'A1:D69', folderUrl, true);
}


/**２次元配列を文字列化する */
function test_convertText(){
  const values = [
    ['John', 'Doe'],
    ['Alice', 'Smith'],
    ['Bob', 'Johnson'],
  ];

  const string = values.map(record => record[0]).join('\n');
  console.log(string);
  console.log(`型：${typeof string}`);
}



/**連番付与とファイル名のリネームが出来るスクリプト*/
function test_renameAllFiles(){
  const folderUrl = 'https://drive.google.com/drive/folders/1z365po-hxBmUNg7MJXD1LEWBYSvqsi9y';
  renameAllFiles(folderUrl);
}




/** for of文とforEach文の違いをテスト */
function test_forOf(){
  const values = [
    ['John', 'Doe'],
    ['Alice', 'Smith'],
    ['Bob', 'Johnson'],
  ];

  for(const [fileName, fileId] of values){
    console.log(fileName, fileId);
  }
}



/**連番作成のテスト */
function test_serialNumbersArray(){
  const serialNumbersArray = Array.from({ length: 999 }, (_, i) => ('00' + (i + 1)).slice(-3));
  console.log(serialNumbersArray);
}



/**フォルダのURLからIDを取得 */
function test_getFolderId(){
  const folderUrl = 'https://drive.google.com/drive/folders/1z365po-hxBmUNg7MJXD1LEWBYSvqsi9y';
  getFolderId(folderUrl);
}



/**ヘッダー行を取得 */
function test_getHeaderRow(){
  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=0';
  const values = getValues(url);
  getHeaderRow(values, 'setRules()');
}



/**採用業務を多用する、リンク付きの名前をHTML形式で表示する */
function test_generateNameWithUrl(){
  const url  = 'https://docs.google.com/spreadsheets/d/1JfPF1KQss6nMA4fHyGvNyAVnaE4zGG80aVx3qKhx4Ow/edit#gid=3542835';
  const html = generateNameWithUrl(url, 0, {name: '氏名', url: 'URL'}, '東京都');
  showHtmlSentence(html, 'HTMLを表示する');
}



function test_convertValuesToObjects(){
  const url      = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=476791012';
  const values   = getValues(url);
  const keys     = ['id', 'name', 'department'];
  convertValuesToObjects(values, 0, keys);
}


function test_replaceStringWithSheetValue(){
  const values      = [["John", "https://example.com"],["Tim", "https://example2.com"]];
  const template    = "こんにちは、{name}さん。URLは{url}です。";

  // nepia_infinity.replaceStringWithSheetValue(template, values, '{name}', '{url}');
  replaceStringWithSheetValue(template, values, '{name}', '{url}');
}


function test_getHeadersRow(){
  const url    = 'https://docs.google.com/spreadsheets/d/1JfPF1KQss6nMA4fHyGvNyAVnaE4zGG80aVx3qKhx4Ow/edit#gid=1404123830';
  const values = getValues(url);
  getHeadersRow(values, 'URL');
}
