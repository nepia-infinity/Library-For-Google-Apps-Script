
/**
 * 開始値から終了値までの整数を生成する
 * 
 * @param  {number} start - 開始値
 * @param  {number} end - 終了値
 * @return {Array.<number>}
 * 　
 */
function generateNumbers(start, end) {

  const array = Array.from({ length: end}, (_, index) => start + index);
  console.info('generateNumbers()を実行中');
  console.info('08_otherに記載');
  console.log(array);

  return array
}




/**
 * ランダムな値を返す
 * @param  {number} end - 最大数値
 * @return {number} number - ランダムな値
 * 
 */
function getRandomNumber(end) {
  console.info('getRandomNumber()を実行中');
  console.info('08_otherに記載');

  const random = Math.floor(Math.random() * end) + 1;
  console.log(`出力されたランダムな整数： ${random}`);

  return random;
}



/**
 * 1次元配列の中で1番大きな数値を返す
 * 
 * @param  {Array.<number>} array - 数値が格納された1次元配列
 * @return {number} 
 */
function findLargestNumber(array){
  console.info('findLargestNumber()を実行中');
  console.info('08_otherに記載');

  const result = array.reduce((accumulator, currentValue) => {
    if(accumulator > currentValue){
      //accumulatorに格納された値が、現在値よりも大きい
      console.log(`accumulator: ${accumulator} > currentValue: ${currentValue}`);
      return accumulator;

    } else {
      //accumulatorに格納された値が、現在値よりも小さい
      //console.log(`accumulator: ${accumulator} < currentValue: ${currentValue}`)
      return currentValue;
      
    }
  });

  console.log(`配列内の最大数値は、${result}です`);
  return result
}




/**
 * ハイフンなしの携帯電話番号を、ハイフンありで返す関数
 * 
 * @param  {string | number} original - 携帯電話の文字列 '09012345678'
 * @param  {string} log - 省略可 - 実行中の関数名を表示する
 * @return {string} ハイフンありの携帯電話 '090-1234-5678'
 * 
 */
function convertCellPhoneNumber(original, log) {

  if(log){
    console.info(`getCellPhoneNumber()を実行中`);
    console.info('08_otherに記載');
  }

  // 数値型だったら文字列化し、文字列だったら引数をそのまま使用する
  const originalString = typeof original === 'number' ? String(original) : original;
  let cellPhoneNumber;

  if (originalString.length < 11) {
    cellPhoneNumber = `0${sliceStringNumber_(originalString, 2)}`;
  } else {
    cellPhoneNumber = sliceStringNumber_(originalString, 3);
  }

  console.warn(`成形後：${cellPhoneNumber}`);
  return cellPhoneNumber
}

/**
 * 携帯電話番号の文字列を作成する
 * 
 * @param  {string} string - 携帯電話の文字列を生成する
 * @param  {number} quantity - 文字をスライスする量
 * @return {string}
 * 
 */
function sliceStringNumber_(string, quantity, log) {
  if(log){
    console.info(`generateCellPhoneNumber_()を実行中`);
    console.info('08_otherに記載');
  }
  
  console.log(`成形前：${string}, string.length: ${string.length}`);

  const parts = [];
  parts.push(string.slice(0, quantity)); // 最初の quantity 桁
  parts.push(string.slice(quantity, quantity + 4)); // 次の 4 桁
  parts.push(string.slice(quantity + 4)); // 残りの部分

  console.log(parts);

  const cellPhoneNumber = parts.join('-');
  return cellPhoneNumber;
}




/**
 * 半角英数字の置換用2次元配列を作成する
 * 
 * 
 */
function generateTwoByteRegularExpression(){
  const twoByteCharacter  = '０１２３４５６７８９';
  const halfSizeCharacter = '0123456789';
  const list = twoByteCharacter.split('');

  console.info(`generateTwoByteRegularExpression()を実行中`);
  console.info('08_otherに記載');

  const lists = list.map(
    (string, i) => [new RegExp(string, 'g'), halfSizeCharacter[i]]
  );

  const newValues = generateTwoByteAlphabetValues_();
  lists.push(...newValues);

  console.log(lists);
  return lists
}




/**
 * 全角のアルファベットと半角のアルファベットの置換用の2次元配列を作成
 * 
 * 
 */
function generateTwoByteAlphabetValues_(){
  const twoByteAlphabet  = Array.from({length: 26}, (_, i) => String.fromCharCode('ａ'.charCodeAt(0) + i));
  const halfSizeAlphabet = Array.from({length: 26}, (_, i) => String.fromCharCode('a'.charCodeAt(0) + i));

  // console.log(twoByteAlphabet);
  // console.log(halfSizeAlphabet);

  const lists = twoByteAlphabet.map(
    (alphabet, i) => [new RegExp(alphabet, 'g'), halfSizeAlphabet[i]]
  );

  //console.log(lists);
  return lists;
}



