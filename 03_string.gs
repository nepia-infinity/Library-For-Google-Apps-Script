/**
 * 捜査対象のテキストから、文字列を消去して欲しい文字列を取得するための関数
 * 使用回数が多くなる傾向にあるので、意図的にconsole.infoを記載していない
 * 
 * @param  {string} string - 操作対象のテキスト
 * @param  {string} reg -  正規表現
 * @param  {Array.<string>} string - 残余引数で何個でも指定可、　置換対象の文字列
 * @return {string}
 * 
 */
function extractText(string, reg, ...params){

  const result   = string.match(reg);
  const replaced = result !== null ? params.reduce((accumulator, current) => accumulator.replace(current, ''), result[0])
    : (console.warn(`${reg}に一致しませんでした`), string);

  console.log(`オリジナルの文字列： ${string}`);
  console.log(`抽出された文字列：  ${replaced}`);

  return replaced

}



/**
 * 氏名からスペースの前の苗字を取得する
 * 
 * @param  {string} fullName - 氏名
 * @param  {boolean} hasLog - 実行中の関数名を表示するかどうか
 * @return {string} 苗字
 * 
 */
function getLastName(fullName, hasLog){

  if(hasLog){
    console.info(`getLastName()を実行中`);
    console.info('03_stringに記載');
  }

  const spaceIndex = fullName.indexOf(' ');

  if (spaceIndex !== -1) {
    // 苗字と名前の前後に空白が含まれていた場合
    const lastName = fullName.slice(0, spaceIndex);
    console.log(`氏名: ${fullName}, 苗字: ${lastName}`);
    return lastName;
  }

  const threeCharLastName = getShortenedLastName(fullName);
  console.log(`氏名: ${fullName}, 苗字: ${threeCharLastName}`);
  return threeCharLastName;
}



/**
 * 3文字の苗字を取得する
 * 
 * @param  {string} fullName - 氏名
 * @param  {boolean} hasLog - 省略可　実行中の関数名を表示するかどうか
 * @return {string} 3文字の苗字
 * 
 */
function getShortenedLastName(fullName, hasLog) {

  if (hasLog) {
    console.info(`getShortenedLastName()を実行中`);
    console.info('03_stringに記載');
  }

  const url    = 'https://docs.google.com/spreadsheets/d/1ng3FcOMax4lbDhqg11UTYHvp6uILLdUFb4_yttI7cy0/edit#gid=692870033';
  const sheet  = SpreadsheetApp.openByUrl(url).getSheetByName('苗字3文字リスト');
  const values = sheet.getDataRange().getValues();

  //苗字リストから3文字を切り取って該当するかどうか検索する
  const lastNameArray = values.map(record => record[0]).filter(value => value);
  const isThreeCharacters = fullName.slice(0, 3);

  // 検索結果がヒットしたらリストから苗字を返す
  const result   = lastNameArray.indexOf(isThreeCharacters);
  const lastName = result !== -1 ? lastNameArray[result] : fullName.slice(0, 2);

  console.log(`苗字：　${lastName}`);
  return lastName;
}



/**
 * 月次連番を作成し、配列で返す
 * 
 * @param  {string} prefix - 接頭辞
 * @param  {number} maxMonths - 生成する月の最大数 (例) 12
 * @param  {string} suffix - 接尾辞
 * @return {Array.<string>}
 * 
 */
function createMonthlyDataNames(prefix, maxMonths, suffix) {

  console.info(`createMonthlyDataNames()を実行中`);
  console.info('03_stringに記載');

  const list = Array.from({ length: maxMonths }, (_, i) => {
    const month = ('0' + (i + 1)).slice(-2);
    return `${prefix}${month}${suffix}`;
  });

  console.log(list);
  return list

}



/**
 * 指定した日付を文字列で返す
 * 
 * @param  {number} offset - 今日を起点に何日後かを進めるか、　-1の場合、昨日の日付になる
 * @param  {string} format - yyyy/MM/dd HH:mmなど日付のフォーマット
 * @return {string}
 * 
 */
function generateFormattedDate(offset, format){
  const date = new Date();
  date.setDate(date.getDate() + offset);
  
  const formattedDate = !format ? Utilities.formatDate(date, 'JST', 'yyyy/MM/dd'):
  Utilities.formatDate(date, 'JST', format)

  console.log(`対象日の日付：${formattedDate}`);
  return formattedDate

}