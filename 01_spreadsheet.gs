/**
 * SpreadsheetのURLからsheetオブジェクトを取得する。
 * シート名が取得したい場合は、2番目の引数に、'sheetName' と指定する
 * トリガー設定可能 getSheets()を使用しているため、実行時間が掛かる。
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
 * FIXME: アクティブなシートを元に処理をするため、トリガー設定は不向き
 * 
 * @param  {string} targetSheetUrl - スプレッドシートのURL
 * @param  {string} string - 引数の省略可　'sheetName' と指定する
 * @return {SpreadsheetApp.Sheet|string} オブジェクトかシート名を返す。
 * 
 */
function getActiveSheetByUrl(targetSheetUrl, string) {

  console.info(`getActiveSheetByUrl()を実行中`);
  console.info(`01_spreadsheetに記載`);

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
 * @param  {SpreadsheetApp.Sheet} sheet - シートオブジェクト
 * @param  {Object.<number>|string} info - 取得開始行と取得開始列 {row: 1, column: 2}　もしくは 'A2:F4' のように指定する
 * @return {SpreadsheetApp.Range} 
 * 
 */
function getRange(sheet, info) {

  console.info('getRange()を実行中');
  console.info('01_spreadsheetに記載');

  if (info && typeof info !== 'string') {

    // info.row が falsy（例: undefined, null, 0, false など）の場合ゼロが設定されます。
    const offset = {
      row: info.row ? info.row - 1 : 0,
      column: info.column ? info.column - 1 : 0,
    };

    const numRows    = sheet.getLastRow() - offset.row;
    const numColumns = sheet.getLastColumn() - offset.column;

    // info.row と info.column が両方指定されていない場合、デフォルトで A1 セルからデータを取得
    const range = sheet.getRange(info.row || 1, info.column || 1, numRows, numColumns);

    console.log(`startRow: ${info.row || 1}, startColumn: ${info.column || 1}, numRows: ${numRows}, numColumns: ${numColumns}`);
    console.log(`取得範囲：　${range.getA1Notation()}`);
    return range;

  } else {
    // infoが'A2:E5'のように文字列だった場合
    const range = sheet.getRange(info);
    console.log(`取得範囲：　${range.getA1Notation()}`);
    return range;
  }
}



/**
 * 指定した列の文字が入力されている最終行を取得する
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
  const arrayWithBlank = values.map(row => row[columnIndex]);
  let lastRow = arrayWithBlank.length;

  // 値があった段階で処理を中断する
  for(let i = arrayWithBlank.length - 1; i >= 0; i--){
    if(arrayWithBlank[i]) break;
    lastRow--;
  }

  console.log('最後の行:', lastRow);
  return lastRow

}



/**
 * 2次元配列の特定の列のみを抽出する
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

  const index = values.findIndex(row => row.includes(query));

  if (index !== -1) {
    const row = index + 1;
    console.log(`ヘッダー行：　${row}`);
    return row;
  }
  console.warn(`一致する行が見つかりませんでした。1を返します`);
  return 1;
}



/**
 * 2次元配列から見出し行の位置を連想配列として取得する
 * FIXME: 見出しの名前が変更になった時の対応が難しい
 * 類似する関数として、generateHeaderIndexがある
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} rowIndex - 見出し行の位置をindexで指定　1行目の場合は0を指定
 * @param  {Object.<string>} targetColumn - 例 {id: 'ID', name: '名前', branch: '所属先'};
 * @return {Object.<number>} 
 */
function replaceHeaderValues(values, rowIndex, targetColumn) {

  console.info(`replaceHeaderValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const column = {};
  const header = values[rowIndex];
  console.log(header);

  // 空のcolumnにプロパティを挿入する
  for(const [key, value] of Object.entries(targetColumn)){
    column[key] = header.indexOf(value);
  }

  console.log(column);
  return column

}



/**
 * 
 * 見出し行の位置を特定する
 * FIXME: 見出し行の項目が全て英語ではないと使えない
 * 類似する関数として、replaceHeaderValuesがある
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} rowIndex - ヘッダー行の位置を指定
 * @param  {Array.<string>} keys - オブジェクトのkeyを格納した1次元配列 (例)　['name', 'url']
 * @return {Object.<number>} 
 */
function generateHeaderIndex(values, rowIndex, keys){
  
  console.info(`generateHeaderIndex()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const header = keys || values[rowIndex];
  const object = Object.fromEntries(
    header.map((value, index) => [value, index])
  );

  console.log(header);
  console.log(object);

  return object;
}



/**
 * 2次元配列内の1次元配列を全てオブジェクトに変換するスクリプト　Googleフォームの回答などにおすすめ
 * 
 * NOTES: https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Object/fromEntries
 * Object.fromEntries() メソッドは、キーと値の組み合わせの配列をオブジェクトに変換する。
 * 
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} columnIndex - 空白をチェックする列のインデックス（0から始まる）
 * @param  {Array.<string>} keys - オブジェクトのkeyを格納した1次元配列 (例)　['name', 'url']
 * @return {Array.<Object.<string|number>>}
 * 
 */

function convertValuesToObjects(values, columnIndex, keys) {

  console.info(`convertValuesToObjects()を実行中`);
  console.info(`01_spreadsheetに記載`);

  // headers にvalues[0], recordsに、valuesの内容をコピーする　（スプレッド構文）
  const [headers, ...records] = values; 
  console.log(values);

  // keysが指定されていればそれをヘッダーとして使用し、そうでなければvaluesの1行目をヘッダーとする
  const customHeaders   = keys || headers;
  const filteredRecords = records.filter(record => record[columnIndex]);

  // 2次元配列内の1次元配列をオブジェクトに置き換える
  // customHeaders[0] = name;
  // [Bob, 20, ramen] -> [name, Bob] -> {name: Bob}
  const objects = filteredRecords.map(record => Object.fromEntries(
    record.map((value, i) => [customHeaders[i], value])
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
  
  const ui = (SpreadsheetApp.getUi() || DocumentApp.getUi());
  const response = ui.prompt(title, sample, ui.ButtonSet.OK);
  const input    = response.getResponseText();

  if (response.getSelectedButton() === ui.Button.OK) {
    console.log(`入力内容：${input}`);
    
  }else{
    console.log('処理が中断されました。');

  }
  return input;
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
function setValues(sheet, info, values, alert) {
  console.info('setValues()を実行中');
  console.info('01_spreadsheetに記載');

  const range = sheet.getRange(info.row, info.column, values.length, values[0].length);
  console.log(`転記範囲：${range.getA1Notation()}`);

  if(alert){
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(`転記範囲に間違いはありませんか？\n\n
      シート名：　${sheet.getName()}
      転記範囲：　${range.getA1Notation()}`, ui.ButtonSet.YES_NO
    );

    if(response === ui.Button.NO){
      console.log('“いいえ”　のボタンが押されました。');
      ui.alert('処理が中断されました。');
      return;
    }
  }

  range.setValues(values);
  console.log('転記が完了しました。');
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
function createTextFinder(url, query, sheetName) {
  const spreadsheet = SpreadsheetApp.openByUrl(url);
  let finder;

  if(sheetName){
    // シート名が指定されている場合、指定したシートでテキスト検索を実行
    const sheet = spreadsheet.getSheetByName(sheetName);
    finder = sheet.createTextFinder(query).useRegularExpression(true);

  }else{
    // シート名が指定されていない場合、全てのシートでテキスト検索を実行
    finder = spreadsheet.createTextFinder(query).useRegularExpression(true);

  }

  const results  = finder.findAll();
  const newArray = results.map(result => ({
    query,
    sheetName: result.getSheet().getName(),
    row:       result.getRow(),
    column:    result.getColumn(),
    range:     result.getA1Notation(),
    value:     result.getValue()
  }));

  console.log(`検索語句：　${query} , 検索結果：　${newArray.length} 件`);
  console.log(newArray);

  return newArray;
}



/**
 * getDataRange()などで取得した2次元配列から必要な列だけを抽出し、新しい2次元配列を作成する
 * 
 * @param  {Array.<Array.<string|number>>} values - 元の2次元配列
 * @param  {Object.<number>} column - 見出し行のオブジェクト (例) {id: 0, name: 1, university: 3}
 * @param  {Array.<string>} queries - 2次元配列から情報を取捨選択するためのキーワード、残余引数なので、いくつでも指定可
 * @return {Array.<Array.<string|number>>} 新しい配列
 *
 */
function selectColumns(values, column, ...queries) {
  console.log(`selectColumns()を実行中`);
  console.log(`01_spreadsheetに記載`);
  
  //{id: 0, name: 1, university: 3}　-> [0, 1, 3]
  const columnsToSelect = Object.values(column);
  console.log(columnsToSelect);

  // 指定された列のインデックスを抽出し、新しい2次元配列を作成する
  const newValues = values.map(row => columnsToSelect.map(index => row[index]));

  // 指定したすべての単語に合致する行のみを残す
  const filtered = newValues.filter(row => {
    return queries.every(query => row.join(',').includes(query));
  });

  console.log(filtered);
  return filtered;
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

  const object = Object.keys(original).reduce((accumulator, key) => (
    {...accumulator, [key]: original[key] += 1} 
  ),{});

  console.log(original);
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

  const result  = [];
  const newData = [...newValues]; // スプレッド演算子を使って、newValuesをコピー
  newData.shift();

  newData.forEach(row => {
    if (existingRecords.indexOf(row[columnIndex]) === -1) {
      console.log(`${newValue[columnIndex]} は新しい値です。転記対象です。`);
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
 * @return {string} 生成されたHTML文字列
 * 
 */
function generateNameWithUrl(url, headerIndex, headers, ...params) {
  console.log(`generateNameWithUrl()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getDisplayValues();
  
  // ヘッダー行を除いたデータをコピー
  const data = [...values];
  data.splice(headerIndex, 1);

  const header      = values[headerIndex];
  const newValues   = Object.entries(headers).map(([key, columnName]) => [key, header.indexOf(columnName)]); // 2次元配列化
  const columnIndex = Object.fromEntries(newValues);  // オブジェクト化

  console.log(header);
  console.log(columnIndex);

  const filtered = data.filter(row => params.every(param => row.includes(param)));
  console.log(filtered);
  console.log(`該当件数：　${filtered.length} 件`);

  // HTMLを生成
  const listItems = filtered.map(row => {
    const name = getLastName(row[columnIndex.name]);
    const link = row[columnIndex.url];
    return `<li><a href="${link}">${name}さん</a></li>`;
  });

  const html = `<ol>${listItems.join('')}</ol>`;
  console.log(html);
  return html;
}


  
/**
 * 2次元配列を縦1列に変換する
 * @param {Array.<Array.<string|number>>} original - 変換対象の元の2次元配列。
 * @returns {Array.<Array.<string|number>>} - 各要素が個別の配列内に収められた新しい2次元配列。
 */
function convertToSingleColumn(original){

  const newValues = original.flat().reduce((accumulator, current) => {
    accumulator.push([current]);
    return accumulator;
  }, []);

  console.log(`convertToSingleColumn()を実行中`);
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
  
  const values = getImageFiles_(folderUrl); // 04_drive.gsに記載
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
 * 
 * 
 * 1. replacePlaceholders() シートの値でプレイスホルダを差し替える
 * 2. makeCopyFile() テンプレートファイルを複製して複製後のファイルIDを返す
 * 3. insertNewSheet() 新しいシートを作成する
 * 
 */
function duplicateReplacedTemplate(info){

  console.info(`duplicateReplacedTemplate()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const values  = getValues(info.sheetUrl);
  const headers = values[0];
  
  // (例) {id: 'ID', name: '名前', address: '住所'}
  // 2次元配列を必要な値のみに取捨選択する
  const column   = buildObjectFromPairs(headers, info.keys, info.headerNames);
  const selected = selectColumns(values, column);
 
  // Google DocumentのURLからファイルIDを抽出
  info['templateFileId'] = getFileId(info.templateFileUrl);
  info['folderId']       = getFolderId(info.folderUrl);

  // テンプレートにシートの値を流し込む
  const newValues = replacePlaceholders_(selected, info);
  const sheet     = insertNewSheet('差込文書');
  sheet.activate();

  // 新しいシートに複製した文書のURLなどを転記する
  const targetRow = getLastRowWithText(sheet.getDataRange().getValues(), 0) + 1;
  console.log(`転記対象行： ${targetRow}`);
  setValues(sheet, {row: targetRow, column: 1}, newValues);

  // ヘッダー行を除いてカウントする
  SpreadsheetApp.getUi().alert(`${newValues.length -1}　件の書類を作成しました`);
}



/**
 * 文書内のplaceholderをシートの値で差込する
 * 複製、差込が完了したファイル名、ファイルID、ファイルURLなどを２次元配列で返す
 * 
 * @param  {Array.<Array.<string>>} values - 2次元配列
 * @param  {Object.<string>} info - ファイル名などを含むオブジェクト
 * @return {Array.<Array.<string>>}
 * 
 */
function replacePlaceholders_(values, info){

  console.info(`replacePlaceholders_()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const [header, ...records] = values;
  const newValues = [['ファイル名', 'ファイルID', 'URL']];

  // ヘッダー行を除いてループ
  for(const record of records){
    // シートの値を使用した置換リスト
    const lists = record.map((value, index) => [`{${info.keys[index]}}`, value]);
    console.log(lists);

    // シートの値でテンプレートのファイル名を置換する
    const replacedFileName = lists.reduce((accumulator, list) => accumulator.replace(...list), info.fileName);
    const duplicatedFileId = makeCopyFile(info.templateFileId, info.folderId, replacedFileName);
    const generatedUrl     = `https://docs.google.com/document/d/${duplicatedFileId}/edit`;
    const targetDocument   = DocumentApp.openById(duplicatedFileId);
    
    console.log(`ファイル名：　${replacedFileName}`);
    console.log(`生成されたURL:　${generatedUrl}`);

    // テンプレートのプレイスホルダーを置き換える
    lists.reduce((accumulator, list) => accumulator.replaceText(...list), targetDocument.getBody());
    newValues.push([replacedFileName, duplicatedFileId, generatedUrl]);
  }
  console.log(newValues);
  return newValues
}



/**
 * 
 * テンプレートを複製してファイルIDを返す
 * 
 * @param  {string} templateFileId - テンプレートのファイルID
 * @param  {string} folderIdentifier - フォルダのURL or フォルダID
 * @param  {string} replacedFileName - 複製したファイルに付ける名前
 * @param  {string} log - 省略可、引数を定義すると、実行中の関数名を表示する
 * @return {string} 複製したファイルのID
 * 
 */
function makeCopyFile(templateFileId, folderIdentifier, replacedFileName, log){

  if(log){
    console.info(`makeCopyFile()を実行中`);
    console.info(`01_spreadsheetに記載`);
  }
  
 // ファイル名、保存場所
  const template = DriveApp.getFileById(templateFileId);
  const folderId = folderIdentifier.includes('https://') ? getFolderId(folderIdentifier) : folderIdentifier;
  const folder   = DriveApp.getFolderById(folderId);

  // makeCopy(ファイル名、保存場所)
  const duplicatedFile   = template.makeCopy(replacedFileName, folder);
  const duplicatedFileId = duplicatedFile.getId();

  return duplicatedFileId
}



/**
 * シート名を引数として渡して新しいシートを渡す。シートが存在する場合は既存のシートオブジェクトを渡す
 * 
 * @param  {string} sheetName - シート名
 * @return {SpreadsheetApp.Sheet} シートオブジェクト
 * 
 */
function insertNewSheet(sheetName){
  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();

  try{
    // 新規作成したシートに引数に指定された名前を渡す
    newSheet.setName(sheetName);
    console.log(`${sheetName}という名前のシートが、新しく作成されました`);

    return newSheet

  }catch{
    // シートがすでに存在していた場合は、シートオブジェクトを返す
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(newSheet);

    console.warn(`引数に指定していたシート名は既に存在していました`);
    console.log(`既存のシート名：${sheet.getName()}`);

    return sheet
  }
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
  console.log(`変更後`);
  console.log(rotated);

  return rotated

}



/**
 * ペアとなるkeyとvalueの配列のペアから、欲しい列のみのcolumnIndexを取得する関数
 * 
 * @param {Array.<string>} header - ヘッダー行、1次元配列
 * @param {Array.<string>} keys - オブジェクトのキーとなる文字列が格納された配列 ['id', 'name', 'url']
 * @param {Array.<string>} array - オブジェクトのキーとなる文字列が格納された配列 ['ID', '名前', 'URL']
 * @return {Object.<string>}
 */
function buildObjectFromPairs(header, keys, array) {

  console.log(`buildObjectFromArray()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const object = keys.reduce((accumulator, key, index) => {
    const result = header.indexOf(array[index]);
    result !== -1 ? accumulator[key] = result : false;
    return accumulator;
  }, {});

  console.log(object);
  return object
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

  sheets.forEach(sheet => {
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
  sheetNames.forEach(sheetName => spreadsheet.insertSheet(sheetName));
  SpreadsheetApp.getUi().alert(`${sheetNames.length}　件のシートを作成しました`);
}



/**
 * オブジェクト内の値を昇順に並び替え、valueをindexで置き換える
 * selectColumnsと併用する事を前提に作成された関数
 * 
 *{id: 0, name: 1, department: 6, birthDate: 2} -> {id: 0, name: 1, birthDate: 2, department: 3}
 *
 * @param  {Object.<number>} object - 値を並び替えて置き換える対象のオブジェクト
 * @param  {number} rowIndex - 比較する列のindex 1列目の場合、0
 * @return {Object.<number>} 値が昇順のインデックスで置き換えられたオブジェクト
 */
function swapWithIndex(object, rowIndex){
  
  console.log('元のオブジェクト');
  console.log(object);

  // オブジェクトを一旦、2次元配列化して昇順で並び替える
  // {id: 0, name: 1, department: 2}　-> [['id', 0],['name', 1],['department', 2]]
  const entries   = Object.entries(object).sort((previous, current) => previous[rowIndex] - current[rowIndex]);

  // keyとvalueを分割代入で取り出し、valueをindexに置き換える
  const replaced  = entries.map(([key, _], index) => [key, index]);
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