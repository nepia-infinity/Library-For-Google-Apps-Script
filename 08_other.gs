
/**
 * 開始値から終了値までの整数を生成する
 * 
 * @param  {number} start - 開始値
 * @param  {number} end - 終了値
 * @return {Array.<number>}
 * 　
 */
function generateNumbers(start, end) {

  console.info('generateNumbers()を実行中');
  console.info('08_otherに記載');

  let newArray = [];
  for(let i = start; i <= end; i++){
    newArray.push(i);
  }
  console.log(newArray);
  return newArray
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
 * @param  {string} string - 携帯電話の文字列 '09012345678'
 * @return {string} string - ハイフンありの携帯電話　 '090-1234-5678'
 * 
 */
function getCellPhoneNumber(string){

  console.info(`getCellPhoneNumber()を実行中`);
  console.info('08_otherに記載');
  console.log(`成形前：${string}`);

  let cellPhoneNumber;

  if(string.match(/[0-9].*/)!== null && string.match(/0.0/) !== null){

    //ゼロ落ちがない場合
    cellPhoneNumber = generateCellPhoneNumber_(string, 3);
    
  }else if(string[0] !== 0){

    console.warn(`ゼロ落ちしている可能性があります。`);
    cellPhoneNumber = `0${generateCellPhoneNumber_(string, 2)}`;

  }

  console.log(`成形後：${cellPhoneNumber}`);
  return cellPhoneNumber
}

/**
 * 携帯電話番号の文字列を作成する
 * 
 * @param  {string} string - 携帯電話の文字列を生成する
 * @param  {quantity} quantity - 文字をスライスする量
 * @return {string}
 * 
 */
function generateCellPhoneNumber_(string, quantity){

  console.info(`generateCellPhoneNumber_()を実行中`);
  console.info('08_otherに記載');

  const first  = string.slice(0, quantity);
  const second = string.replace(first, '').slice(0, 4);

  //1番最後の4桁のみを残す
  const third  = string
  .replace(first, '')
  .replace(second, '');

  const cellPhoneNumber = `${first}-${second}-${third}`;
  return cellPhoneNumber
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



