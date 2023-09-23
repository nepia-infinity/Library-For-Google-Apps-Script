/**
 * SpreadsheetのURLからsheetオブジェクトを取得する。
 * シート名を取得したい場合は、2番目の引数に、'sheetName' と指定する
 * トリガー設定可能 getSheets()を使用しているため、実行時間が掛かる。
 * 
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @param  {string} string - 引数の省略可。'sheetName' と指定する
 * @return {SpreadsheetApp.Sheet|string} オブジェクトかシート名を返す。
 * 
 */
function getSheetByUrl(sheetUrl, string) {

  console.info(`getSheetByUrl()を実行中`);
  console.info(`01_spreadsheetに記載`);
  
  const spreadsheet    = SpreadsheetApp.openByUrl(sheetUrl);
  const sheets         = spreadsheet.getSheets();
  const sheetInfoArray = sheetUrl.split('#gid=');

  console.log(sheetInfoArray);

  //シートIDを、文字列から数値に変換する
  const sheetId = Number(sheetInfoArray[1]);

  for(const sheet of sheets){
    if(sheet.getSheetId() === sheetId && !string){
      console.log(`sheetId: ${sheetId} typeof: ${typeof sheetId}`);
      console.log(`sheetName: ${sheet.getName()}`);
      return sheet

    }else if(sheet.getSheetId() === sheetId && string === 'sheetName'){
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
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @param  {string} string - 引数の省略可　'sheetName' と指定する
 * @return {SpreadsheetApp.Sheet|string} オブジェクトかシート名を返す。
 * 
 */
function getActiveSheetByUrl(sheetUrl, string) {

  console.info(`getActiveSheetByUrl()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const activeSheet    = SpreadsheetApp.getActiveSheet();
  const sheetInfoArray = sheetUrl.split('#gid='); //['https....', 'sheetId(typeof string)'];

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

    // info.row が falsy（例: undefined, null, 0, false など）の場合ゼロを設定
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
  for(let i = lastRow - 1; i >= 0; i--){
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

  console.info(`generateArray()を実行中`);
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
function getHeadersRow(values, query){

  console.info(`getHeadersRow()を実行中`);
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
 * 引数で渡されたオブジェクトの値をindexOfの結果に差し替える
 * 類似する関数として、generateHeadersIndexがある
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} rowIndex - 見出し行の位置をindexで指定　1行目の場合は0を指定
 * @param  {Object.<string>} targetColumn - 例 {id: 'ID', name: '名前', branch: '所属先'};
 * @return {Object.<number>} {id: 0, name: 1, branch: 2};
 */
function replaceHeadersValues(values, rowIndex, targetColumn) {

  console.info(`replaceHeadersValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const header = values[rowIndex];
  console.log(header);

  // targetColumnの値をindexOfの結果で差し替える (例) {id: 'ID', name: '名前', branch: '所属先'};
  const newValues = Object.entries(targetColumn).map(([key, value]) => [key, header.indexOf(value)]);
  const column    = Object.fromEntries(newValues);

  console.log(column);
  return column

}



/**
 * 
 * 見出し行の位置を特定する
 * 類似する関数として、replaceHeadersValuesがある
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} rowIndex - ヘッダー行の位置を指定
 * @param  {Array.<string>} keys - オブジェクトのkeyを格納した1次元配列 (例)　['name', 'url']
 * @return {Object.<number>} 
 */
function generateHeadersIndex(values, rowIndex, keys){
  
  console.info(`generateHeadersIndex()を実行中`);
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
 * 2次元配列内の1次元配列を全てオブジェクトに変換するスクリプト
 * FIXME: 見出しが日本語だとオブジェクトの各keyも日本語になってしまう
 * 
 * @param  {Array.<Array.<string|number>>} values - 2次元配列
 * @param  {number} columnIndex - 空白をチェックする列のインデックス（0から始まる）
 * @param  {Array.<string>} keys - 省略可、オブジェクトのkeyを格納した1次元配列 (例)　['name', 'url']
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
  
  const ui = (SpreadsheetApp.getUi() || DocumentApp.getUi() || SlidesApp.getUi());
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
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @return {Array.<Array.<string|number>>} values - 2次元配列
 * 
 */
function getValues(sheetUrl) {

  console.info(`getValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const sheet  = getSheetByUrl(sheetUrl);
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
 * @param  {boolean} hasAlert - setValuesを実行する前にアラートを表示するかどうか
 * @return {SpreadsheetApp.Range} 
 * 
 */
function setValues(sheet, info, values, hasAlert) {

  console.info('setValues()を実行中');
  console.info('01_spreadsheetに記載');

  const range = sheet.getRange(info.row, info.column, values.length, values[0].length);
  console.log(`転記範囲：${range.getA1Notation()}`);

  if(hasAlert){
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
 * @param  {string} params - 取り出したいオブジェクトのkeyをいくつでも指定可
 * @retrun {Array.<Object.<srting | number>>}
 * 
 */
function reduceObjectKeys(values, ...params){

  console.info(`reduceObjectKeys()を実行中`);
  console.info(`01_spreadsheetに記載`);

  console.log(`valuesから　${params}　の${params.length}つを取り出す`);
  console.log(values);

  const reduced = values.reduce((accumulator, current) => {
    const obj = {};
    params.map(param => {
      obj[`${param}`] = current[param];

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
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @param  {string} query - スプレッドシートで検索したい単語
 * @param  {string} sheetName - 検索したいシートの名前　検索対象を絞りたい場合に使用
 * @return {Array.<Object.<srting | number>>} 
 */
function createTextFinder(sheetUrl, query, sheetName){

  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
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

  console.info(`selectColumns()を実行中`);
  console.info(`01_spreadsheetに記載`);
  
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

  console.info(`modifyObject()を実行中`);
  console.ifno(`01_spreadsheetに記載`);

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

  // スプレッド演算子を使って、ヘッダー行を除いてnewValuesをコピー
  const newData = [_, ...newValues];

  // シートに記載のIDなどの1次元配列と比較して、一致しないIDを転記対象とみなす
  const results = newData.map(row => {
    if (existingRecords.indexOf(row[columnIndex]) === -1) {
      console.warn(`${newValue[columnIndex]} は新しい値です。転記対象です。`);
      return [row];
    }
  });

  console.log(results);
  return results;
}



/**
 * 
 * URL付きテキストの生成する 管理表に記載されている応募者名のリンク付きURLを生成する事を想定している
 * 残余引数については下記のページを参照
 * https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Functions/rest_parameters
 * 
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @param  {number} rowIndex - 見出し行の配列番号
 * @param  {Object.<string>} columnNames - 見出し行に使用されている項目名をオブジェクトで指定　　（例）{name: '氏名', url: 'URL'}
 * @param  {string} params - 検索クエリ複数可　（例）　active,　合格など残余引数として指定できる
 * @return {string} 生成されたHTML文字列
 * 
 */
function generateNameWithUrl(sheetUrl, rowIndex, columnNames, ...params) {

  console.info(`generateNameWithUrl()を実行中`);
  console.info(`01_spreadsheetに記載`);

  // ヘッダー行を削除する
  const values   = getValues(sheetUrl);
  const filtered = getFilteredValues(values, params);
  console.log(filtered);
  console.log(`該当件数：　${filtered.length} 件`);
  
  // 引数で渡されたオブジェクトの値をindexOfの結果に差し替える
  const column = replaceHeadersValues(values, rowIndex, columnNames);
  console.log(column);

  // HTMLを生成
  const listItems = filtered.map(row => {
    const name = getLastName(row[column.name]);
    const link = row[column.url];
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

  console.info(`convertToSingleColumn()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const newValues = original.flat().reduce((accumulator, current) => {
    accumulator.push([current]);
    return accumulator;
  }, []);

  console.warn(`変換後`);
  console.log(newValues);

  return newValues
}



/**
 *  文字列内の指定されたプレイスホルダーをスプレッドシートの値で置き換え2次元配列で返す
 * 
 * @param  {string} template - 置換対象となる文章  (例)　'こんにちは、{name}さん。URLは{url}です。'
 * @param  {Array.<Array.<string>>} values - 主にシートから取得した2次元配列 (例)　 [['John', 'https://example.com'],['Tim', 'https://example2.com']]
 * @param  {Array.<string>} params - 置換対象となる単語、プレイスホルダーの文字列 (例)　'{name}', '{url}'
 * @return {Array.<Array.<string>>}
 * 
 */
function replaceStringWithSheetValue(originalText, values, ...params) {

  console.info(`replaceStringWithSheetValue()を実行中`);
  console.info(`01_spreadsheetに記載`);

  //シートの2次元配列を取り出し、originalText内のプレイスホルダーを書き換える
  const newValues = values.map(row => {
    const replaced = params.reduce((accumulator, current, index) => {
      const regex = new RegExp(current, 'g');
      return accumulator.replace(regex, row[index]);
    }, originalText);
    return [replaced];
  });

  console.log(newValues);
  return newValues;
  
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
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @param  {string} query - 検索する文字列
 * @param  {number} queryColumnIndex - 照合したい列
 * @param  {number} resultColumnIndex - データを取得したい列
 * @return {string} 取得したいデータ 
 * 
 */
function findDataByQuery(sheetUrl, query, queryColumnIndex, resultColumnIndex) {

  console.info(`findDataByQuery()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const values   = getValues(sheetUrl);
  const foundRow = values.find(row => row[queryColumnIndex] === query);
  const result   = foundRow ? foundRow[resultColumnIndex] : null;
  
  console.log(`query: ${query} result: ${result}`);
  (result === null || result === undefined) ? console.log("データが見つかりませんでした") : null;

  return result;
}



/**
 * 複数の列の値を1つのセルに結合します。
 * 
 * @param  {number} sheetUrl - スプレッドシートのURL
 * @return {Array.<Array><string>>} 2次元配列
 * 
 */
function combineColumnToSingleCell(sheetUrl){

  console.log(`combineColumnToSingleCell()を実行中`);
  console.log(`01_spreadsheetに記載`);

  const values    = getValues(sheetUrl);
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
  
  console.info(`addImageSheet()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const values   = getImageFiles_(folderUrl); // 04_drive.gsに記載
  let sheet      = SpreadsheetApp.getActiveSheet();
  const response = Browser.msgBox(`${sheet.getName()}に画像を挿入します。よろしいでしょうか？`, Browser.Buttons.OK_CANCEL);
  console.log(`ダイアログの選択肢：${response}`); //ok, cancel

  // アクティブなシートで処理を実行していいかアラートを出す。ダメな場合はシート名を入力させる
  response === 'ok' ? true : sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(showPrompt('シート名を入力してください', 'シート1'));

  let targetRow = startRow;

  for(const [fileName, fileId] of values){

    const image = createImageFromBlob_(fileId, fileName);
    const range = sheet.getRange(targetRow, startColumn)
    range.setValue(image);
    console.log(`処理対象範囲: ${range.getA1Notation()}`);

    targetRow += 1;

  }

  SpreadsheetApp.getUi().alert(`${values.length} 件の画像を挿入しました`);

}



/**
 * 与えられたファイルIDとファイル名から画像を生成し、新しい画像オブジェクトを作成して返す関数です。
 *
 * @param {string} fileId - 画像ファイルのGoogle Drive上のファイルID
 * @param {string} fileName - 画像のファイル名
 * @param {boolean} hasLog - 実行中の関数名を表示する
 * @return {SpreadsheetApp.Sheet} - 画像オブジェクト
 */
function createImageFromBlob_(fileId, fileName, hasLog){

  if(hasLog){
    console.info(`createImageFromBlob_を実行中`)
    console.info(`01_spreadsheetに記載`);
  }

  const imageBlob   = DriveApp.getFileById(fileId).getBlob();
  const contentType = imageBlob.getContentType();
  const base64      = Utilities.base64Encode(imageBlob.getBytes());
  const imageString = "data:" + contentType + ";base64, " + base64;

  const image = SpreadsheetApp.newCellImage()
    .setSourceUrl(imageString)
    .setAltTextTitle(fileName)
    .setAltTextDescription("-")
    .build();

  return image;
}



/**
 * 差込文書作成を自動化する関数、テンプレートとなるドキュメント、保存先等を指定する必要がある
 * 
 * @param {Object.<string>} info - オブジェクト
 * sheetUrl, fileName, keys, headerNames, templateUrl, folderUrlを定義する必要がある
 * オブジェクトの定義例については、下記URLを参照
 * https://note.com/nepia_infinity/n/n4320954a3851#21dbed4d-7fa0-4cc2-ba3f-b0bb7a819a65
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
 * @param  {boolean} hasLog - 省略可、引数を定義すると、実行中の関数名を表示する
 * @return {string} 複製したファイルのID
 * 
 */
function makeCopyFile(templateFileId, folderIdentifier, replacedFileName, hasLog){

  if(hasLog){
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

  console.info(`insertNewSheet()を実行中`);
  console.info(`01_spreadsheetに記載`);

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

  console.info(`setRules()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const values = getValues(sourceSheetUrl);
  const array  = generateArray(values, columnIndex);
  const sheet  = SpreadsheetApp.getActiveSheet();

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

  console.info(`rotateValues()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const rotated = values[0].map((_, i) => values.map(row => row[i]));
  console.log(`変更後`);
  console.log(rotated);

  return rotated

}



/**
 * ペアとなるkeyとvalueの配列のペアから、欲しい列のみのcolumnIndexを取得する関数
 * 
 * @param {Array.<string>} headers - ヘッダー行、1次元配列
 * @param {Array.<string>} keys - オブジェクトのキーとなる文字列が格納された配列 ['id', 'name', 'url']
 * @param {Array.<string>} array - オブジェクトのキーとなる文字列が格納された配列 ['ID', '名前', 'URL']
 * @return {Object.<string>}
 */
function buildObjectFromPairs(headers, keys, array) {

  console.info(`buildObjectFromPairs()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const object = keys.reduce((accumulator, key, index) => {
    const result = headers.indexOf(array[index]);
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

  console.info(`deleteSpecificSheets()を実行中`);
  console.info(`01_spreadsheetに記載`);

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

  console.info(`generateMultipleSheets()を実行中`);
  console.info(`01_spreadsheetに記載`);

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

  console.info(`swapWithIndex()を実行中`);
  console.info(`01_spreadsheetに記載`);
  
  console.log('元のオブジェクト');
  console.log(object);

  // オブジェクトを一旦、2次元配列化して昇順で並び替える
  // {id: 0, name: 1, department: 2}　-> [['id', 0],['name', 1],['department', 2]]
  const entries = Object.entries(object).sort((previous, current) => previous[rowIndex] - current[rowIndex]);

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
function getItemCounts(array, items){

  console.info(`getItemCounts()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const counts = {};
  array.forEach(item => {
    if (items.includes(item)) {
      counts[item] = (counts[item] || 0) + 1;
    }
  });

  console.log(counts);
  return counts;
}



/**
 * スプレッドシートの指定範囲から2次元配列を取得し、名前 -> メールアドレスのように一致する別の値に変換し、新たな2次元配列を作成
 * [['A', 'B', 'C'], ['D', 'E', 'F']] -> [['A, B, C'], ['D, E, F']]
 * 
 * 類似する関数としてconvertToSingleColumn()がある
 * 
 * @param  {Object.<string | number>} info - オブジェクト 
 * オブジェクトの指定例 { sheetUrl:  'https://...', stringRange: 'A2:D45', queryColumnIndex: 0,  resultColumnIndex: 1 }
 * @param  {Array.<string>} additionalInfo - 例えば、メールアドレスなど
 * @return {Array.<Array.<string>>}
 * 
 */
function convertSheetDataToQueryResults(info, ...additionalInfo){

  console.info(`convertSheetDataToQueryResults()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const sheet  = getActiveSheetByUrl(info.sheetUrl);
  const range  = getRange(sheet, info.stringRange);
  const values = range.getDisplayValues();
  console.log(values);

  const newValues = values.map(row => {
    // （例） 名前 -> メールアドレス
    const array = row.map(query => findDataByQuery(info.sheetUrl, query, info.queryColumnIndex, info.resultColumnIndex));

    // 引数が複数、つまり配列ならば、forEachを使用して一つずつ値を取り出して追加
    (additionalInfo && typeof additionalInfo === 'object') ? additionalInfo.forEach(newValue => array.unshift(newValue))
    : array.unshift(additionalInfo)

    return [array.join(',')];
  });

  console.log(newValues);
  return result

}



/**
 * 住所からアパート名を抽出し、列を分けた新しい配列を作成する
 * 必要に応じて、insertNewSheetやsetValuesと組み合わせるといいかも
 * FIXME: 虎ノ門など、鎌ヶ谷など、一部の漢字がカタカナと誤判定されてしまうケースがある
 * 
 * @param  {string} url - スプレッドシートのURL
 * @param  {number} rowIndex - ヘッダー行の位置（1行目の場合は0と指定）
 * @param  {number} columnIndex - 住所が記載されている列の位置（1列目の場合は0と指定）
 * @return {Array.<Array.<string>>}
 * 
 */
function splitAddressColumn(sheetUrl, rowIndex, columnIndex){

  console.info(`splitAddressColumn()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const values = getValues(sheetUrl);
  values.splice(rowIndex, 1);

  // 全角数字を半角数字に直す変換リストを生成
  const twoBiteConvertLists = generateTwoByteRegularExpression();
  let lists = [
    [/.*(市|区|町|村)/, ''],
    [/[0-9]{1,4}-[0-9]{1,4}-[0-9]{1,4}/, '']
  ];

  // 既存の変換リストと結合させる
  lists = lists.concat(twoBiteConvertLists);
  console.log(lists);

  // 2次元配列から住所のみの1次元配列を作成
  const addressArray = generateArray(values, columnIndex);

  // 住所からアパート名を抽出する
  const newValues = addressArray.map(original =>{
    const result = original.match(/[ァ-ンヴー].*|[A-Za-z].*/);
    const apartmentName = (result !== null)
    ? lists.reduce((accumulator, current) => accumulator.replace(...current), result[0])
    : '';

    // アパート名を抜いた住所
    const address = (apartmentName) ? original.replace(apartmentName, '') : original
    return [original, address, apartmentName];
  });

  console.log(newValues);
  return newValues
}



/**
 * 金融機関コード（4桁） or 支店名コード（3桁）になるように '0'で字詰めする
 * 
 * @param  {string} sheetUrl - スプレッドシートのURL
 * @param  {number} columnIndex - 金融機関コードが記されている列
 * @param  {boolean} isBankCode - 金融機関コードかどうか、falseの場合は支店名コード
 * @return {Array.<Array.<string>>} 
 * 
 */
function formatBankCode(sheetUrl, columnIndex, isBankCode){

  console.info(`formatBankCode()を実行中`);
  console.info(`01_spreadsheetに記載`);

  const values = getValues(sheetUrl);
  const array  = generateArray(values, columnIndex);

  // 金融機関コード4桁 or 支店名コードの3桁になるように0で字詰めをする
  const targetLength    = (isBankCode) ? 4 : 3;
  const convertBankCode = (number => String(number).padStart(targetLength, '0'));

  const formatedArray = array.map(code => {
    const formatedCode = (typeof code === 'number') ? convertBankCode(code) : code
    return [formatedCode];
  });

  console.log(formatedArray);
  return formatedArray

}