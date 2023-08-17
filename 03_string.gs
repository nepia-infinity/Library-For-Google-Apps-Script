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
    return ''
  }
}




/**
 * 
 * スペースの前の苗字を返す。'久保 翼'だとしたら、久保を返す
 * @param  {string} string - 氏名
 * @return {string} 苗字
 * 
 */
function getLastName(string){
  let lastName;
  const result = string.match(/.*\s/);

  console.log(`getLastName()を実行中`);
  console.info('03_stringに記載');
  
  console.log(result);
  console.log(`氏名:　${string}`);

  if(result !== null){
    //苗字と名前の前後に空白が含まれていた場合
    lastName = result[0].replace(/\s/, '');

  }else if(3 <= string.length ){
    console.warn(`3文字以上の文字列です。苗字リストを検索して該当したら3文字の苗字を返します。`);
    lastName = getThreeCharLastName_(string);

  }else{
    lastName = string;
    
  }
  console.log(`苗字：　${lastName}`);
  return lastName;
}


/**
 * 3文字の苗字の場合の処理
 * 
 * @param  {string} string - 氏名
 * @param  {string} lastName - 苗字
 * 
 */
function getThreeCharLastName_(string){
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('苗字3文字リスト');
  const values = sheet.getDataRange().getValues();

  const lastNameArray = values.map(record => record[0]).filter(value => value);
  const temp          = string.slice(0, 3);
  const result        = lastNameArray.indexOf(temp);
  console.log(`indexOfの結果：　${result}`);

  let lastName;

  if(result !== -1){
    lastName    = lastNameArray[result];
    const row   = result + 1; 
    const range = getRange(sheet, `B${row}`);

    //検索でヒットした回数を更新する
    const previous = range.getValue();
    const current  = previous + 1;
    range.setValue(current);
    
  } else{
    lastName = string.slice(0, 2);
    
  }
  return lastName
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
  const monthlyDataNames = [];

  for (let i = 1; i <= maxMonths; i++) {
    const month = ('0' + i).slice(-2);
    const dataName = `${prefix}${month}${suffix}`;
    monthlyDataNames.push([dataName]);
  }

  console.log(monthlyDataNames);
  return monthlyDataNames;
}


