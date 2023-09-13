/**
 * SpreadsheetのURLからsheetオブジェクトを取得する。
 * シート名が取得したい場合は、2番目の引数に、'sheetName' と指定する
 * トリガー設定可能 getSheets()を使用しているため、実行時間が掛かる。
 * FIXME: sheet.getLastRow(), sheet.getDataRange().getValues()でエラーが生じる？
 * 
 * @param  {string} url - スプレッドシートのURL
 * @param  {string} keyWord - 引数の省略可。'sheetName' と指定する
 * @return {SpreadsheetApp.Sheet|string} オブジェクトかシート名を返す。
 * 
 */
function getSheetByUrl(url, keyWord) {
  const spreadsheet    = SpreadsheetApp.openByUrl(url);
  const sheets         = spreadsheet.getSheets();
  const sheetInfoArray = url.split('#gid=');

  console.log(`getSheetByUrl()を実行中`);
  console.log(sheetInfoArray);

  //シートIDを、文字列から数値に変換する
  const sheetId = Number(sheetInfoArray[1]);

  for(const sheet of sheets){
    if(sheet.getSheetId() === sheetId && !keyWord){
      console.log(`sheetId: ${sheetId} typeof: ${typeof sheetId}`);
      console.log(`sheetName: ${sheet.getName()}`);
      return sheet

    }else if(sheet.getSheetId() === sheetId && keyWord === 'sheetName'){
      const sheetName = sheet.getName();
      console.log(`sheetName: ${sheetName} typeof: ${typeof sheetId}`);
      return sheetName

    }
  }
}






/**
 * SpreadsheetのURLからsheetオブジェクトを取得する。
 * シート名を取得したい場合は、2番目の引数に、'sheetName' と指定する
 * アクティブなシートを元に処理をするため、トリガー設定は不向き
 * 
 * 最終更新日　2023/09/13
 * 
 * @param  {string} targetSheetUrl - スプレッドシートのURL
 * @param  {string} string - 引数の省略可　'sheetName' と指定する
 * @return {SpreadsheetApp.Sheet|string} オブジェクトかシート名を返す。
 * 
 */
function getActiveSheetByUrl(targetSheetUrl, string) {
  const activeSheet    = SpreadsheetApp.getActiveSheet();
  const sheetInfoArray = targetSheetUrl.split('#gid='); //['https....', 'sheetId(typeof string)'];

  console.log(sheetInfoArray);

  // 前述のsheetIdが、型も含めて完全一致しない場合は処理を終了する
  if(activeSheet.getSheetId() !== Number(sheetInfoArray[1])){
    console.log(`シート名：${activeSheet.getName()}`);
    console.warn(`処理対象のシートではないため、処理を終了します`);
    return

  }else if(string === 'sheetName'){
    const sheetName = activeSheet.getName();
    console.warn(`シート名：${sheetName} 型：${typeof sheetName}`);
    return sheetName;

  }else{
    console.log(`シート名：${activeSheet.getName()} 型：${typeof activeSheet}`);
    return activeSheet
  }
}




/**
 * 指定したシートの範囲を取得する。
 * 
 * 最終更新日　2023/04/28
 * 
 * @param  {SpreadsheetApp.Sheet} sheet - シートオブジェクト
 * @param  {Object.<number>|string} info - 取得開始行と取得開始列 {row: 1, column: 2}　もしくは 'A2:F4' のように指定する
 * @return {SpreadsheetApp.Range} 
 * 
 */
function getRange(sheet, info){
  let range, offset, numRows, numColumns;

  console.info(`getRange()を実行中`);
  console.info(`01_spreadsheetに記載`);
  console.log(info);

  if(info && typeof info !== 'string'){
    // infoがオブジェクトだった場合
    offset  = info.row - 1;
    numRows = sheet.getLastRow() - offset;

    if(info.column !== 1){
      offset     = info.column - 1;
      numColumns = sheet.getLastColumn() - offset;
      range      = sheet.getRange(info.row, info.column, numRows, numColumns);

    }else{
      // info.columnが1の時
      numColumns = sheet.getLastColumn();
      range      = sheet.getRange(info.row, info.column, numRows, numColumns);

    }

    console.log(`startRow: ${info.row}, startColumn: ${info.column}, numRows: ${numRows}, numColumns: ${numColumns}`);

  }else{
    // infoが'A2:E5'のように文字列だった場合
    range = sheet.getRange(info);
    
  }
  console.log(`取得範囲：　${range.getA1Notation()}`);
  return range
}



/**
 * 指定した列の文字が入力されている最終行を取得する
 * 
 * 最終更新日　2023/06/16
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} columnIndex - 2列目が欲しい場合は1と指定
 * @return {number}
 * 
 */
function getLastRowWithText(values, columnIndex){

  console.info('getLastRowWithText()を実行中');
  console.info('01_spreadsheetに記載');

  // 途中の空白行を含む1次元配列を生成する
  const generateArrayWithBlank = values.map(row => row[columnIndex]);
  console.log(generateArrayWithBlank);

  let lastRow    = generateArrayWithBlank.length;
  const reversed = generateArrayWithBlank.reverse();

  for(const value of reversed){
    if(!value){
      //　空白行だったら、maxLengthから1を引く
      lastRow -=1;

    }else if(value){
      break;
    }
  }//for
  console.log(`最終行：　${lastRow}`);
  return lastRow
}



/**
 * 2次元配列の特定の列のみを抽出する
 * 
 * 最終更新日　2023/04/28
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} columnIndex - 数字で指定、0始まりなので2列目の場合は1と指定
 * @return {Array.<string|number>} 1次元配列
 */
function generateArray(values, columnIndex){

  console.info(`generateArray_()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const array = values.map(record => record[columnIndex]).filter(value => value);
  console.log(array);
  return array;
}



/**
 * シートオブジェクトを引数にアクティブなセルの値、行、列などの情報を取得する
 * 
 * 最終更新日　2023/04/28
 * 
 * @param  {SpreadsheetApp.Sheet} sheet - シートオブジェクト
 * @return {Object.<number|string>}
 */
function getActiveCell(sheet){

  console.info(`getActiveCell()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const activeCell   = sheet.getActiveCell();
  const activeDetail = {
    sheetName: activeCell.getSheet().getName(),
    row:       activeCell.getRow(),
    column:    activeCell.getColumn(),
    value:     activeCell.getValue(),
    range:     activeCell.getA1Notation()
  }

  console.log(activeDetail);
  return activeDetail
}



/**
 * 空白行のあるシートから見出し行を見つけるスクリプト
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {string} query - 見出し行に使用されている単語
 * @return {number}
 */
function getHeaderRow(values, query){
  console.info(`getHeaderRow()を実行中`);
  console.info(`01_spreadsheetに記載`);

  let row = 1;

  for(let i = 0; i < values.length; i++){
    if(values[i].indexOf(query) !== -1){
      row += i;
      console.log(`ヘッダー行：　${row}`);
      return row
    }
  }//for
}



/**
 * 2次元配列から見出し行の位置を連想配列として取得する
 * FIXME: 見出しの名前が変更になった時の対応が難しい
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} rowIndex - 見出し行の位置をindexで指定　1行目の場合は0を指定
 * @param  {Object.<string>} targetColumn - 例 {id: 'ID', name: '名前', branch: '所属先'};
 * @return {Object.<number>} 
 */
function replaceHeaderValues(values, rowIndex, targetColumn) {

  console.info(`replaceHeaderValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const header = values[rowIndex];
  console.log('header');
  console.log(header);

  // key, valueを一人ずつ取り出す
  const entries = Object.entries(targetColumn);
  const column  = {};

  console.log('Object.entriesの結果');
  console.log(entries);

  // 空のcolumnにプロパティを挿入する
  for(const [key, value] of entries){
    column[key] = header.indexOf(value);

  }

  console.log(column);
  return column

}



/**
 * 
 * 見出し行の位置を特定する
 * FIXME: 見出し行の項目が全て英語ではないと使えない
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @return {Object.<number>} 
 */
function generateHeaderIndex(values){
  
  console.log(`generateHeaderIndex_()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const header = values.shift();//1次元配列
  const object = new Map();

  header.map((value, index) => object[value] = index);
  console.log(object);

  return object
}



/**
 * 2次元配列内の1次元配列を全てオブジェクトに変換するスクリプト　Googleフォームの回答などにおすすめ
 * FIXME: ヘッダー行が1行目にない非構造化データの処理には向かない
 * 
 * NOTES: https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Object/fromEntries
 * Object.fromEntries() メソッドは、キーと値の組み合わせのリストをオブジェクトに変換します。
 * 
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @return {Array.<Object.<string|number>>}
 * 
 */

function convertValuesToObjects(values) {

  console.log(`generateHeaderIndex_()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const [headers, ...records] = values; // headers にvalues[0], recordsにそれ以外が代入される
  console.log(values);

  // 2次元配列内の1次元配列をオブジェクトに置き換える
  // headers[0] = name;
  // [Bob, 20, ramen] -> [name, Bob] -> {name: Bob}
  const objects = records.map(record => Object.fromEntries(
    record.map((value, i) => [headers[i], value])
  ));

  console.log(objects);
  return objects
}



/**
 * 入力画面を表示させ、入力内容を取得する
 * FIXME: SpreadsheetやDocument以外だとエラーが生じる
 * 
 * @param  {string} title - 表題（例）検索したい単語を入力してください
 * @param  {string} sample - 入力例
 * @return {string}
 */
function showPrompt(title, sample) {

  console.info(`showPrompt()を実行中`);
  console.info(`01_spreadsheetに記載`);
  
  let ui;

  try{
    ui = SpreadsheetApp.getUi();
  }catch{
    ui = DocumentApp.getUi();
  }

  const response = ui.prompt(title, sample, ui.ButtonSet.OK);
  const input    = response.getResponseText();

  switch (response.getSelectedButton()){
    case ui.Button.OK:
      console.log(`入力された内容：${input}`);
      break;

    case ui.Button.CLOSE:
      console.log('閉じるボタンが押されました。');
      break;

    default:
    console.log('処理が中断されました。');

  }//switch

  return input

}

/**
 * シートの表示内容を2次元配列で取得する
 * 
 * @param  {string} url - スプレッドシートのURL
 * @return {Array.<Array.<string|number>>} values - 2次元配列
 * 
 */
function getValues(url) {

  console.info(`getValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getValues();
  console.log(values);
  
  return values;

}//end​




/**
 * 指定したシートの範囲に2次元配列を転記する
 * 
 * 最終更新日　2023/05/01
 * 
 * @param  {SpreadsheetApp.Sheet} sheet - シートオブジェクト
 * @param  {Object.<number>} info - 取得開始行と取得開始列 {row: 1, column: 2}
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {string} alert - setValuesを実行する前にアラートを表示するかいなか　(例)　アラート
 * @return {SpreadsheetApp.Range} 
 * 
 */
function setValues(sheet, info, values, alert){

  console.info('setValues()を実行中');
  console.info('01_spreadsheetに記載');
  console.log(info);

  const range = sheet.getRange(info.row, info.column, values.length, values[0].length);
  console.log(`転記範囲：${range.getA1Notation()}`);

  if(!alert){
    range.setValues(values);

  }else if(alert){
    const ui       = SpreadsheetApp.getUi();
    const response = ui.alert(`転記範囲に間違いはありませんか？\n\n
      シート名：　${sheet.getName()}
      転記範囲：　${range.getA1Notation()}`, ui.ButtonSet.YES_NO
    );

    switch (response){
      case　ui.Button.YES:
        console.log('“はい”　のボタンが押されました。');
        range.setValues(values);
        break;

      case ui.Button.NO:
        console.log('“いいえ”　のボタンが押されました。');
        ui.alert('処理が中断されました。');
        break;

      default:
        console.log('処理が中断されました。');
        return
    }
  }
}



/**
 * 指定した単語のみを含む2次元配列を生成
 * 
 * 最終更新日：2023/06/18
 * 
 * @param  {Array.<Array.<string|number>>} values - スプレッドシートのURL
 * @param  {string|Array.<string>} params - 特定のキーワード、残余引数いくつでも指定可
 * @return {Array.<Array.<string|number>>}
 * 
 */
function getFilteredValues(values, ...params){
  
  console.info(`getFilteredValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const filtered = values.filter(row => params.every(param => row.join(',').includes(param)));
  console.log(params);
  console.log(filtered);
  console.log(`${filtered.length} 項目一致しました`);

  return filtered
  
}




/**
 * オブジェクトの中から引数に指定したkeyのみを取り出す
 * 
 * @param  {Array.<Object.<srting | number>>} values - [{id: 'jp123', name: 'nobita'}, {id: 'jp456', name: 'shizuka'}]
 * @param  {string} theArgs - 取り出したいオブジェクトのkeyをいくつでも指定可
 * @retrun {Array.<Object.<srting | number>>}
 * 
 */
function reduceObjectKeys(values, ...theArgs){

  console.info(`reduceObjectKeys()を実行中`);
  console.info(`01_spreadsheetに記載`);
  console.log(`valuesから　${theArgs}　の${theArgs.length}つを取り出す`);
  console.log(values);

  const reduced = values.reduce((accumulator, current) => {
    const obj = {};
    theArgs.map(arg => {
      obj[`${arg}`] = current[arg];

    });
    console.log(obj);
    
    accumulator.push(obj);
    return accumulator
    
  }, []);

  console.log(reduced);
  return reduced;

}



/**
 * HTMLファイルを表示する
 * 
 * @param  {string} file - HTMLファイル
 * @param  {string} title - 表題
 * @return {string}
 * 
 */
function showHtmlSentence(file, title){

  console.info(`showHtmlSentence()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const ui   = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(file)
  .setWidth(450)
  .setHeight(200);

  ui.showModelessDialog(html, title);
}




/**
 * spliceメソッドを使用時に分かりやすくログを表示する
 * 
 * @param  {Array.<string>} array - 1次元配列
 * @param  {number} start - 処理の開始位置
 * @param  {number} quantity - 変化量　0の時は値を追加、それ以外は置換
 * @param  {string|number} query - 文字列、値いずれも可。2番目の引数によって役割が変化する
 * 
 */
function splice(array, start, quantity, query){

  console.info(`splice()を実行中`);
  console.info(`01_spreadsheetに記載`);
  
  if(!query){
    console.warn(`配列の${start}番目から、値を${quantity}つ削除します。`);
    const deleted = array.splice(start, quantity);
    console.log(`削除された値：${deleted}`);

  }else if(quantity === 0 && query){
    console.warn(`配列の${start}番目に、値、${query}を追加する`);
    array.splice(start, quantity, query);

  }else if(quantity !== 0 && query){
    console.warn(`配列の${start}番目の値${array[start]}を、${query}で置換する`);
    array.splice(start, quantity, query);
  }

  console.log(array);
  return array
}


/**
 * 
 * 1次元配列内の要素の重複を省く
 * 
 * @param  {Array.<string>} array - 1次元配列
 * @return {Array.<string>}
 * 
 */
function removeDuplicates(array){

  console.info(`removeDuplicates()を実行中`);
  console.info(`01_spreadsheetに記載`);
  console.log(array);

  const newArray = Array.from(new Set(array));
  console.log(newArray);
  return newArray
}


/**
 * 検索ワード、行、列、範囲などの情報を連想配列で返す
 * 全てのシートが検索対象
 * 
 * @param  {string} url - スプレッドシートのURL
 * @param  {string} query - スプレッドシートで検索したい単語
 * @param  {string} sheetName - 検索したいシートの名前　検索対象を絞りたい場合に使用
 * @return {Array.<Object.<srting | number>>} 
 */
function createTextFinder(url, query, sheetName){
  const spreadsheet = SpreadsheetApp.openByUrl(url);
  let finder;

  if(sheetName){
    const sheet = spreadsheet.getSheetByName(sheetName);
    finder      = sheet.createTextFinder(query).useRegularExpression(true);

  }else if(!sheetName){
    finder = spreadsheet.createTextFinder(query).useRegularExpression(true);
    
  }

  const results = finder.findAll();
  const length  = results.length;

  console.log(`検索語句：　${query} , 検索結果：　${length} 件`);
  const keys     = ['query', 'sheetName', 'row', 'column', 'range', 'value'];
  let object     = {};
  let temp       = [];
  const newArray = [];

  for(const result of results){
    const info = {
      sheetName: result.getSheet().getName(),
      row:       result.getRow(),
      column:    result.getColumn(),
      range:     result.getA1Notation(),
      value:     result.getValue()
    }

    temp.push(query, info.sheetName, info.row, info.column, info.range, info.value);
    //console.log(temp);

    for(const [i, value] of temp.entries()){
      const key = keys[i];
      object[key] = value;
    }
    newArray.push(object);

    //配列に加えた後で初期化
    object = {};
    temp   = [];
  }
  console.log(newArray);
  return newArray
}



/**
 * getDataRange()などで取得した2次元配列から必要な列だけを抽出し、新しい2次元配列を作成する
 * 
 * @param  {Array.<Array.<string|number>>} values - 元の2次元配列
 * @param  {Object.<number>} column - 見出し行のオブジェクト (例) column = {id: 0, name: 1, university: 3}
 * @param  {string} query - 2次元配列から情報を取捨選択するためのキーワード
 * @return {Array.<Array.<string|number>>} 新しい配列
 *
 */
function selectColumns(values, column, query){

  console.log(`selectColumns()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const numbers = Object.values(column);

  console.log(column);
  console.log(`Object.valuesの実行結果`);
  console.log(numbers);//1次元配列

  //indexに該当する列だけを残して2次元配列を作成する
  const newValues = values.map(array => array.reduce((accumulator, current, index) =>{
      if(numbers.includes(index)){
        accumulator.push(current);
      }
    return accumulator
    }, [])//reduce
  );//map

  //console.log(newValues);
  
  if(!query){
   //queryが省略されており、定義されていない場合、空白行の配列を取り除く
    const filtered = newValues.filter(row => row[0] !== '');
    console.log(filtered);

    return filtered

  }else if(query){
    //newValuesから、さらに特定の単語が含まれている配列のみを残す
    const filtered = newValues.filter(row => row.indexOf(query) !== -1);
    console.log(filtered);

    return filtered
  }
}




/**
 * オブジェクトの中身をsortし、上書きする
 * 
 * @param  {Object.<number>} column - 見出し行の情報　{id: 0, name: 1, department: 2}
 * @param  {number} index - SORTする対象のindex
 * @return {Object.<number>}
 * 
 */
function sortInsideObject(column, index) {

  console.log(`sortInsideObject()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const values = Object.entries(column);
  console.log(`Object.entriesの実行結果`);
  console.log(values);

  // 2次元配列を日付などでSORTする
  const newValues = values.sort((previous, current) => (previous[index] < current[index]) ? -1 : 1);
  console.log(`SORT後の2次元配列`);
  console.log(newValues);

  const keys = newValues.map(record => record[0]).filter(value => value);
  // console.log(keys);

  const object = new Object();
  keys.map((key, index) => object[`${key}`] = index);
  console.log(object);

  return object
}



/**
 * 
 * getRange()メソッドで使うために、オブジェクトの値を全て　+1　にする
 * {id: 0, name: 1, department: 2} => {id: 1, name: 2, department: 3}
 * 
 * @param  {Object.<number>} original - 見出し行の情報　{id: 0, name: 1, department: 2}
 * @return {Object.<number>} 
 * 
 */

function modifyObject(original) {

  console.log(`modifyObject()を実行中`);
  console.log(`01_spreadsheetに記載`);

  console.log(original);

  const object = Object.keys(original).reduce((accumulator, key) => (
    {...accumulator, [key]: original[key] += 1} 
  ),{});

  console.log(object);
  return object
}



/**
 * 
 * 新しい値かどうかを確認し転記対象の場合、配列に加える。
 * 
 * @param  {Array.<string|number>} existingValues - 既に存在する値が格納された1次元配列（例：シートなどに転記済みのID, URLなど）
 * @param  {Array.<Array.<string|number>>} newValues - 全ての値が格納された2次元配列（例：Gmailなどから取得したデータ）
 * @param  {number} columnIndex - newValues内の列インデックス（例：IDやURLなどを取り出す列番号）
 * @return {Array.<Array.<string|number>>} - 新たに転記する対象となる値を含む2次元配列
 * 
 */

function selectNewValues(existingRecords, newValues, columnIndex){

  console.log(`selectNewValues() 関数を実行中`);
  console.log(`01_spreadsheet に記載中`);

  newValues.shift();

  console.log('newValues ヘッダー行削除後');
  console.log(newValues);

  let result = [];

  newValues.map(row => {
    if(existingRecords.indexOf(row[columnIndex]) === -1){
      console.log(`${row[columnIndex]} は新しい値です。転記対象です。`);
      result.push(row);
    }
  });

  console.log(result);
  return result;
}



/**
 * 
 * URL付きテキストの生成する 管理表に記載されている応募者名のリンク付きURLを生成する事を想定している
 * 残余引数については下記のページを参照
 * https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Functions/rest_parameters
 * 
 * @param  {string} url - スプレッドシートのURL
 * @param  {number} headerIndex - 見出し行の配列番号
 * @param  {Object.<string>} object - 見出し行に使用されている項目名をオブジェクトで指定　　（例）{name: '氏名', url: 'URL'}
 * @param  {string} params - 検索クエリ複数可　（例）　active,　合格など残余引数として指定できる
 * @return {string} 
 * 
 */
function generateNameWithUrl(url, headerIndex, object, ...params){

  console.log(`generateNameWithUrl()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[headerIndex];
  const column = {
    name: header.indexOf(object.name),
    url:  header.indexOf(object.url)
  }
  console.log(header);
  console.log(column);

  //　URLが含まれるものだけを残す everyメソッドは配列内全ての条件を満たすとtrueで返す
  let filtered = values.filter(row => params.every(param => row.includes(param)));
  console.log(filtered);
  console.log(`該当件数：　${filtered.length}　件`);

  let string = '<ol>';

  // HTMLを生成
  filtered.map(row => string += `<li><a href ="${row[column.url]}">${getLastName(row[column.name])}さん</a></li>`);
  string += '</ol>';

  console.log(string);
  return string

}

  
/**
 * 2次元配列を、各元素を個別の配列要素とする新しい2次元配列に変換します。
 * @param {Array.<Array.<string|number>>} original - 変換対象の元の2次元配列。
 * @returns {Array.<Array.<string|number>>} - 各要素が個別の配列内に収められた新しい2次元配列。
 */
function reformatWithAlternateValues(original){

  const newValues = original.flat().reduce((accumulator, current) => {
    accumulator.push([current]);
    return accumulator;
  }, []);

  console.log(`reformatWithAlternateValues()を実行中`);
  console.log(`01_spreadsheetに記載`);

  console.log(`変換前`);
  console.log(original);

  console.log(`変換後`);
  console.log(newValues);

  return newValues
}



/**
 * テンプレートの文章にシートの値を流し込むスクリプト
 * 
 * @param {string} template - 置換対象となる文章やテンプレート
 * @param {Array.<number>} targetArray - 置換対象となる単語、プレイスホルダーを格納した1次元配列
 * @param {Array.<Array.<string>>} values - 主にシートから取得した2次元配列
 * 
 */
function replaceTemplateWithValues(template, targetArray, values) {

  console.log(`replaceTemplateWithValues()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const newValues   = [];
  
  for(const row of values){
    const lists = targetArray.map((string, index) => [new RegExp(string, 'g'), row[index]]);
    console.log(lists);

    // 2次元配列を作成しテンプレートの文章を置換する
    // [[/置換対象/g, '差込する値'], [/置換対象/g, '差込する値']]
    const replaced = lists.reduce((acc, [regex, replacement]) => acc.replace(regex, replacement), template);
    console.log(replaced);

    newValues.push([replaced]);

  }

  console.log(newValues);
  return newValues
  
}



/**
 * 配列内のオブジェクトを昇順で並び替える
 * 
 * @param  {Array.<Object.<string|number>>} array - オブジェクトを格納した1次元配列
 * @param  {string} key - オブジェクトのkeyを指定
 * @return {Array.<Object.<string|number>>}
 * 
 */
function sortInsideArray(array, key){

  console.log(`sortInsideArray()を実行中`);
  console.log(`01_spreadsheetに記載`);

  console.log(`sort前`);
  console.log(array);

  //条件が当てはまらない場合、配列の最後に追加し、-1以外だったら位置を指定して追加する
  const sortedArray = array.reduce((acc, current) => {
    const insertIndex = acc.findIndex(item => item[key] > current[key]);
    insertIndex === -1 ? acc.push(current) : acc.splice(insertIndex, 0, current);
    return acc;
  }, []);

  console.warn(`sort後`);
  console.log(sortedArray);
  return sortedArray

}


/**
 * スプレッドシートの2次元配列内のデータを検索し、一致した行かつ指定した列の情報を返す
 * 
 * @param  {number} url - スプレッドシートのURL
 * @param  {number} queryColumnIndex - 照合したい列
 * @param  {string} query - 検索する文字列
 * @param  {number} targetColumnIndex - データを取得したい列
 * @return {string} 取得したいデータ 
 * 
 */
function findDataByQuery(url, queryColumnIndex, query, targetColumnIndex) {

  console.log(`findDataByQuery()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const values   = getValues(url);
  const foundRow = values.find(row => row[queryColumnIndex] === query);
  const result   = foundRow ? foundRow[targetColumnIndex] : null;
  
  console.log(`query: ${query} result: ${result}`);
  (result === null || result === undefined) ? console.log("データが見つかりませんでした") : null;

  return result;
}


/**
 * 複数の列の値を1つのセルに結合します。
 * 
 * @param  {number} url - スプレッドシートのURL
 * @return {Array.<Array><string>>} 2次元配列
 * 
 */
function combineColumnToSingleCell(url){

  console.log(`combineColumnToSingleCell()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const values    = getValues(url);
  const header    = values.shift();
  const newValues = values.map(row => {
    const newRow = row.reduce((acc, value, index) => {
      acc += `${header[index]}:${value}\n`;
      index += 1;
      return acc;
    }, '');
    return [newRow];
  });

  console.log(newValues);
  return newValues
}


/**
 * 
 * スプレッドシートに画像を挿入するスクリプト
 * 
 * @param {string} folderUrl - Google DriveのフォルダURL
 * @param {number} startRow - スプレッドシートに画像を挿入開始する行
 * @param {number} startColumn - スプレッドシートに画像を挿入開始する列
 * 
 */
function addImageToSheet(folderUrl, startRow, startColumn) {

  const values = getImageFiles_(folderUrl); //04_drive.gsに記載
  console.log(`${values.length} 件`);

  console.log(`addImageSheet()を実行中`);
  console.log(`01_spreadsheetに記載`);

  let sheet      = SpreadsheetApp.getActiveSheet();
  const response = Browser.msgBox(`${sheet.getName()}に画像を挿入します。よろしいでしょうか？`, Browser.Buttons.OK_CANCEL);

  console.log(`ダイアログの選択肢：${response}`); //ok, cancel

  // アクティブなシートで処理を実行していいかアラートを出す。ダメな場合はシート名を入力させる
  response === 'ok' ? true : sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(showPrompt('シート名を入力してください', 'シート1'));

  let targetRow = startRow;

  for(const [fileName, fileId] of values){

    const imageBlob   = DriveApp.getFileById(fileId).getBlob();
    const contentType = imageBlob.getContentType();
    const base64      = Utilities.base64Encode(imageBlob.getBytes());
    const imageStr    = "data:" + contentType + ";base64, " + base64;

    const image = SpreadsheetApp.newCellImage()
    .setSourceUrl(imageStr)
    .setAltTextTitle(fileName)
    .setAltTextDescription("-")
    .build();

    const range = sheet.getRange(targetRow, startColumn)
    range.setValue(image);
    console.log(`処理対象範囲: ${range.getA1Notation()}`);

    targetRow += 1;

  }

  SpreadsheetApp.getUi().alert(`${values.length} 件の画像を挿入しました`);

}



/**
 * 他のシートからデータ入力規則を参照出来る
 * 
 * @param {string} sourceSheetUrl - データ入力規則を参照したいシート
 * @param {number} columnIndex - 列
 * @param {string} cell - 範囲の文字列 A1表記　（例） A2:Cなど
 * 
 */
function setRules(sourceSheetUrl, columnIndex, cell){

  const values = getValues(sourceSheetUrl);
  const array  = generateArray(values, columnIndex);
  const sheet  = SpreadsheetApp.getActiveSheet();

  console.log(`setRules()を実行中`);
  console.log(`01_spreadsheetに記載`);

  sheet.getRange(cell)
  .setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInList(array, true)
  .build());

}


/**
 * 
 * 2次元配列を回転させる
 * 
 * @param  {Array.<Array.<string>>} values - 2次元配列
 * @return {Array.<Array.<string>>}
 * 
 */
function rotateValues(values) {

  console.log(`rotateValues()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const rotated = values[0].map((_, i) => values.map(row => row[i]));
  console.log(`変更前`);
  console.log(values);

  console.log(`変更後`);
  console.log(rotated);

  return rotated

}



/**
 * 対となる配列を用意して、それぞれkeyとvalueを格納したオブジェクトを作成する
 * 
 * @param  {Array.<string>} header - スプレッドシートのヘッダー行 ex: values[0], values.shift()
 * @param  {Array.<string>} keys - オブジェクトのkeyを格納した1次元配列
 * @param  {Array.<string>} values - オブジェクトのvalueを格納した1次元配列
 * @return {Object.<number>}
 * 
 */
function buildObjectFromPairs(header, keys, values){

  console.log(`buildObjectFromPairs()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const result = keys.reduce((object, key, index) => {
    object[key] = header.indexOf(values[index]);
    return object;
  }, {});

  console.log(result);
  return result
  
}



/**
 * 配列からオブジェクトを構築する関数
 * FIXME: 余計な列がある場合は不具合が起きかねないので使用不可
 * 
 * @param {Array.<string>} header - オブジェクトの値となる文字列が格納された配列 ['ID', '氏名', 'URL']
 * @param {Array.<string>} keys - オブジェクトのキーとなる文字列が格納された配列 ['id', 'name', 'url']
 * @return {Object.<string>} - 構築されたオブジェクト
 */
function buildObjectFromArray(header, keys) {

  console.log(`buildObjectFromArray()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const result = keys.reduce((object, key, index) => {
      object[key] = header.indexOf(header[index]);
      return object;
  }, {});

  console.log(result);
  return result
}



/**
 * テンプレートの文書にシートの値を差込し、複製するスクリプト
 * 複製後、URLなどを書き出す　PDF化には、43. convertDocToPdf()が便利
 * 
 * @param {Object.<string|Array.<string>>} info - オブジェクトに以下のkeyが必要　　　sheetUrl, fileName, keys, headerName, templateUrl, folderUrl
 * 
 * 
 */
function duplicateReplacedTemplate(info){

  console.log(`duplicateReplacedTemplate()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const values  = getValues(info.sheetUrl);
  const headers = values[0];
  
  // 2次元配列を必要な値のみに取捨選択する
  const column   = buildObjectFromPairs(headers, info.keys, info.headerNames);
  const selected = selectColumns(values, column);
  const folderId = getFolderId(info.folderUrl);

  // Google DocumentのURLからファイルIDを抽出する
  const templateFileId = extractText(info.templateFileUrl, /\/d\/[a-zA-Z0-9_-].*/, '/d/', '/edit');

  // テンプレートにシートの値を流し込み、
  const newValues = replacePlaceholders_(selected, info.keys, info.fileName, templateFileId, folderId);
  const count     = newValues.length -1;
  console.log(`作成された　${count}　件`);

  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  let sheetName  = '差込文書のURL';

  try{
    newSheet.setName(sheetName);

  }catch{
    sheetName = showPrompt('新しいシート名を入力してください', `例：${sheetName}`);
    newSheet.setName(sheetName);

  }

  // 新しいシートに複製した文書のURLなどを転記する
  newSheet.activate();
  newSheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);
  SpreadsheetApp.getUi().alert(`${count}　件の書類を作成しました`);

}
  



/**
 * 2次元配列の値をテンプレートの文書のプレイスホルダーを置換して複製する
 * 
 * @param {Array.<Array.<string>>} values - テンプレートに差込する値
 * @param {Array.<string>} keys - 1次元配列　 ex. ['id', 'name', 'url', 'cellPhone', 'address']
 * @param {string} fileName - 差込文書のファイル名
 * @param {string} templateFileId - テンプレートのファイルID
 * @param {string} folderId - 作成した文書の保存先
 * @return {Array.<Array.<string>>} 2次元配列
 * 
 * 
 */
function replacePlaceholders_(values, keys, fileName, templateFileId, folderId){

  console.log(`replacePlaceholders_()を実行中`);
  console.log(`01_spreadsheetに記載`);
  
  const [header, ...records] = values;
  let newValues = [['ファイル名', 'ファイルID', 'URL']];

  for(const row of records){
    const lists = row.map((value, index) => [`{${keys[index]}}`, value]);
    console.log(lists);

    const replacedFileName = lists.reduce((accumulator, list) => accumulator.replace(...list), fileName);
    console.log(`ファイル名：　${replacedFileName}`);

    // ファイル名、保存場所
    const template = DriveApp.getFileById(templateFileId); 
    const folder   = DriveApp.getFolderById(folderId);

    // makeCopy(ファイル名、保存場所)
    const duplicatedDocument   = template.makeCopy(replacedFileName, folder);
    const duplicatedDocumentId = duplicatedDocument.getId();

    // 生成されたドキュメントのURL
    const generatedUrl   = `https://docs.google.com/document/d/${duplicatedDocumentId}/edit`;
    const targetDocument = DocumentApp.openById(duplicatedDocumentId);

    console.log(`生成されたURL:　${generatedUrl}`);

    // テンプレートのプレイスホルダーを置き換える
    lists.reduce((accumulator, list) => accumulator.replaceText(...list), targetDocument.getBody());
    newValues.push([replacedFileName, duplicatedDocumentId, generatedUrl]);
    
  }
  console.log(newValues);
  return newValues
}



/**
 * 削除したくないシートを指定して、それ以外のシートを全て削除する
 * 
 * @param {Array.<string>} excludedSheetNames - 削除したくないシート名を格納した配列
 * 
 */
function deleteSpecificSheets(excludedSheetNames){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets      = spreadsheet.getSheets();
  let count         = 0;

  sheets.map(sheet => {
    if(excludedSheetNames.indexOf(sheet.getName())=== -1){
      console.log(`処理対象のシート：　${sheet.getName()}`);
      spreadsheet.deleteSheet(sheet);
      count += 1;
    };
  });

  const string = `${count}　件シートを削除しました`;
  console.log(string);

  SpreadsheetApp.getUi().alert(string);

}



/**
 * 配列に指定したシート名に沿ってシートを新規作成する
 * 
 * @param {Array.<string>} sheetNames - 作成したいシート名を格納した配列
 * 
 */
function generateMultipleSheets(sheetNames){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.map(sheetName => spreadsheet.insertSheet(sheetName));
  SpreadsheetApp.getUi().alert(`${sheetNames.length}　件のシートを作成しました`);
}



/**
 * オブジェクト内の値を昇順に並び替え、それに対応するインデックスで置換する
 *
 * @param  {Object.<number>} object - 値を並び替えて置き換える対象のオブジェクト
 * @return {Object.<number>} 値が昇順のインデックスで置き換えられたオブジェクト
 */
function swapWithAscendingIndex(object){
  
  console.log('元のオブジェクト');
  console.log(object);

  const entries = Object.entries(object).sort((a, b) => a[1] - b[1]);
  const replaced = entries.map(([key, _], index) => [key, index]);
  const newObject = Object.fromEntries(replaced);

  console.log(`sortしてindexで置換後`);
  console.log(newObject);

  return newObject

}



/**
 * 配列内の指定された項目の出現回数を数える関数、COUNTIFのような挙動する
 * 配列内の要素が日本語の場合、keyも日本語になる点は注意が必要
 * 文字列に変換する場合、　　　　　Object.keys(counts).map(key => `${key}: ${counts[key]}`).join('\n')
 * 2次元配列に変換する場合、Object.entries(counts)
 * 
 * @param  {Array.<string>} array - 1次元配列
 * @param  {Array.<string>} items - 1次元配列 (例)　[ 'とても満足', 'やや満足', 'どちらともいえない', 'やや不満', '不満' ]
 * @return {Object.<number>} オブジェクト　（例）　{ 'とても満足': 3, 'やや不満': 2, 'やや満足': 3, '不満': 1 }
 * 
 */
function getItemCounts(array, items) {
  const counts = {};

  array.forEach(item => {
    if (items.includes(item)) {
      counts[item] = (counts[item] || 0) + 1;
    }
  });

  console.log(counts);
  return counts;
}