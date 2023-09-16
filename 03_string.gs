/**
 * 捜査対象のテキストから、文字列を消去して欲しい文字列を取得するための関数
 * 
 * @param  {string} text - 操作対象のテキスト
 * @param  {string} reg -  正規表現
 * @param  {string} string - 置換対象の文字列、残余引数で何個でも可
 * @return {string}
 */
function extractText(text, reg, ...params){

  console.info('extractText()を実行中');
  console.info('03_stringに記載');

  let string;

  if(text.match(reg) !== null){
    string = text.match(reg)[0];

    //配列に格納されている置換対象の文字列で置換していく
    for(const targetWord of params){
      string = string.replace(targetWord, '');
    }
   
    console.log(`オリジナルの文字列：　${text}`);
    console.log(`抽出された文字列：　${string}`);
    
    return string

  }else{
    console.log(`matchの結果：${text.match(reg)}`);
    console.warn(text);
    return text
  }
}



/**
 * 氏名からスペースの前の苗字を取得します。
 * @param  {string} fullName - 氏名
 * @return {string} 苗字
 */
function getLastName(fullName, log) {
  if(log){
    console.info(`getLastName()を実行中`);
    console.info('03_stringに記載');
  }
  
  const spaceIndex = fullName.indexOf(' ');

  if (spaceIndex !== -1) {
    // 苗字と名前の前後に空白が含まれていた場合
    const lastName = fullName.slice(0, spaceIndex);
    console.log(`氏名: ${fullName}, 苗字: ${lastName}`);
    return lastName;

  } else if (fullName.length >= 3) {

    const threeCharLastName = getThreeCharLastName(fullName);
    console.log(`氏名: ${fullName}, 苗字: ${threeCharLastName}`);
    return threeCharLastName;

  } else {
    console.log(`氏名: ${fullName}, 苗字: ${fullName}\n`);
    return fullName;

  }
}

/**
 * 3文字の苗字を取得します。
 * @param  {string} fullName - 氏名
 * @return {string} 3文字の苗字
 */
function getThreeCharLastName(fullName) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('苗字3文字リスト');
  const values = sheet.getDataRange().getValues();

  const lastNameArray = values.map(record => record[0]).filter(value => value);
  const temp   = fullName.slice(0, 3);
  const result = lastNameArray.indexOf(temp);

  if (result !== -1) {
    const lastName = lastNameArray[result];
    const row      = result + 1;
    const range    = getRange(sheet, `B${row}`);

    // 検索でヒットした回数を更新する
    const previous = range.getValue();
    const current = previous + 1;
    range.setValue(current);

    return lastName;

  }else{
    const lastName = fullName.slice(0, 2);
    return lastName;
  }
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
  const list = Array.from({ length: maxMonths }, (_, i) => {
    const month = ('0' + (i + 1)).slice(-2);
    return `${prefix}${month}${suffix}`;
  });

  console.log(list);
  return list

}


