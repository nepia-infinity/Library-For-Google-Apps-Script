/**
 * dateオブジェクトを、yyyy/MM/ddなどの指定した文字列で返す。
 * E - 曜日の指定があった際に　wed → 水　のように変換する
 * 
 * @param  {date} date - dateオブジェクト
 * @param  {sting} format - 'yyyy/MM/dd', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd (E)'
 * @param  {boolean} hasLog - 実行中の関数名をログ表示にするかどうか
 * @return {string} （例）2022/04/06
 */
function formatDate(date, format, hasLog){

  if(hasLog){
    console.info(`formatDate()を実行中`);
    console.info(`02_calenderに記載`);
  }

  const formatDate = Utilities.formatDate(date, 'JST', format);
  const isMatch    = formatDate.match(/[a-zA-Z]/) !== null;
  const day        = convertDay(date.getDay());

  // 2022/04/06 (wed) → 2022/04/06 (水)　のように変換
  return isMatch ? formatDate.replace(/[a-zA-Z]{3}/, `${day}`) : formatDate;

}



/**
 * 今日を起点として、対象日までの日付と曜日の2次元配列として返す
 * 
 * @param  {sting} string - '2023/04/29'のように 'yyyy/MM/dd' 形式で指定する
 * @return {Array.<Array.<string>>}
 * 
 */
function generateDateStringValues(string) {

  console.info(`generateDateStringValues()を実行中`);
  console.info(`02_calenderに記載`);

  const targetDate = new Date(string);
  const today      = new Date();

  const isPast   = targetDate < today;
  const isFuture = today < targetDate;

  if (isPast || isFuture) {
    
    const startDate = isPast ? targetDate : today;
    const endDate   = isPast ? today : targetDate;

    console.warn(`startDate: ${formatDate(startDate, 'yyyy/MM/dd')}, endDate: ${formatDate(endDate, 'yyyy/MM/dd')}`);

    const dayCount  = Math.floor((endDate - startDate) / (24 * 60 * 60 * 1000));
    const newValues = Array.from({ length: dayCount + 1 }, (_, i) => {
      const currentDate = new Date(startDate);
      currentDate.setDate(currentDate.getDate() + i);
      return formatDate(currentDate, 'yyyy/MM/dd E').split(' ');
    });

    console.log(newValues);
    return newValues;

  }else{
    // 過去でも未来でもない場合、空の配列を返す
    return [];
  }
}



/**
 * 英語表記の曜日、sat, wedなどを、'日月火水木金土'のいずれかに変換する
 * 
 * @param  {number} tempNumber - date.getDay()
 * @param  {boolean} hasLog - 実行中の関数名をログ表示にするかどうか
 * @return {string} 
 * 
 */
function convertDay(tempNumber, hasLog){

  const dayOfWeek = '日月火水木金土';
  const daysArray = dayOfWeek.split('');
  const day       = daysArray[tempNumber];

  if(hasLog){
    console.info(`convertDay()を実行中`);
    console.info(`02_calenderに記載`);

    console.log(daysArray);
    console.log(`daysArray[${tempNumber}]　${day}曜日`);
  }

  return day

}



/**
 * 引数にした日時からnカ月後の予定を2次元配列で取得する
 * 
 * @param  {string} string - 'yyyy/MM/dd'形式で日付を指定　（例）'2023/04/28'
 * @param  {number} offset - nカ月後にあたる　数値で指定
 * @param  {string} calId - カレンダーID
 * @return {Array.<Array.<string>>} 
 * 
 */
function getCalAllEvents(string, offset, calId) {

  console.info(`getCalAllEvents()を実行中`);
  console.info(`02_calenderに記載`);

  const cal   = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
  const query = showPrompt('検索したい予定名を入力してください', '（例） 定例会議');
        
  const startTime = new Date(string);
  const endTime   = new Date();
  endTime.setMonth(endTime.getMonth() + offset);

  const events = cal.getEvents(startTime, endTime)
  .filter(event => event.getTitle().includes(query))
  .map(event => ({
    title:       event.getTitle(),
    date:        formatDate(event.getStartTime(), 'yyyy/MM/dd'),
    day:         convertDay(event.getStartTime().getDay()),
    start:       formatDate(event.getStartTime(), 'HH:mm'),
    end:         formatDate(event.getEndTime(), 'HH:mm'),
    description: event.getDescription(),
    guests:      guestList_(event.getGuestList(), event.getCreators())
  }));

  console.log(events);

  const keys   = Object.keys(events[0]);
  const values = events.map(event => keys.map(key => event[key]));

  console.warn(`オブジェクトを2次元配列に変換`);
  console.log(values);
  console.log(`該当する予定： ${values.length}件`);

  return values

}



/**
 * カレンダーの出席者を取得する
 * 
 * @param  {Array.<string>} guests - 出席者を格納した配列
 * @param  {Array.<string>} creators - 予定の作成者を格納した配列
 * @return {string} 
 * 
 */
function guestList_(guests, creators) {

  console.info(`guestList_()を実行中`);
  console.info(`02_calenderに記載`);

  const guestEmails = guests.map(guest => guest.getEmail());
  
  if (creators) {
    // 主催者を既存の配列の先頭に追加
    const creator = creators[0];
    guestEmails.unshift(creator);
  }

  // 配列を文字列化する
  const guestsList = guestEmails.join();
  return guestsList;
}



/**
 * シートの値に基づいて予定を登録するスクリプト
 * 3番目の引数、calIdについては省略した場合、自分のカレンダーに予定が登録される
 * 
 * @param {string} sheeturl　- スプレッドシートのURL
 * @param {number} rowIndex - ヘッダー行の位置
 * @param {string} calId - カレンダーID　省略可　 (例) *****@gmail.com
 * 
 */
function registerEvents(sheeturl, rowIndex, calId) {

  console.info(`registerEvents()を実行中`);
  console.info(`02_calenderに記載`);

  const sheet  = getSheetByUrl(sheeturl);
  const values = sheet.getDataRange().getValues();
  const column = assignHeaderKeys_(values, rowIndex);

  // 3番目の引数、calIdが無かった場合、自分のカレンダーIDを使用する
  calId = calId ? calId : CalendarApp.getDefaultCalendar().getId();

  let count = 0;

  values.forEach((row, index) => {
    const eventId   = registerEventIfNotRegistered_(row, column, calId);
    const targetRow = index + 1;

    console.log(`処理対象行:　${targetRow}行目`);

    if (eventId) {
      sheet.getRange(targetRow, column.eventId + 1).setValue(eventId);
      sheet.getRange(targetRow, column.status + 1).setValue('登録済');
      count += 1;
    }
  });
  
  SpreadsheetApp.getUi().alert(`${count}件の予定を登録しました`);
  
}



/**
 * スプレッドシートの登録ステータスに応じて登録処理を進める関数
 * 
 * @param  {Array.<string|date>} row - 1次元配列
 * @param  {Obeject.<number>} column - 列を特定するために必要なオブジェクト
 * @param  {string} calId - カレンダーID
 * @return {string} Meet URL
 * 
 */
function registerEventIfNotRegistered_(row, column, calId) {

  console.info(`registerEventIfNotRegistered_()を実行中`);
  console.info(`02_calenderに記載`);

  // イベントID or 登録ステータスが空白の場合のみ登録処理を実行する
  if (row[column.eventId] === '' && row[column.status] === '') {

    // イベントの開始時刻
    const startTime = new Date(row[column.date]);
    startTime.setHours(row[column.startTime].getHours());
    startTime.setMinutes(row[column.startTime].getMinutes());

    // イベントの終了時刻
    const endTime = new Date(startTime);
    endTime.setHours(row[column.endTime].getHours());
    endTime.setMinutes(row[column.endTime].getMinutes());

    const eventObject = {
      calId:       calId,
      title:       row[column.title],
      startTime:   Utilities.formatDate(startTime, 'JST', "yyyy-MM-dd'T'HH:mm:ss.000+09:00"),
      endTime:     Utilities.formatDate(endTime, 'JST', "yyyy-MM-dd'T'HH:mm:ss.000+09:00"),
      description: row[column.description],
      attendees:   generateAttendees_(row[column.attendees]),
    }
    return createEventWithMeetUrl_(eventObject);
  }
  return null;
}



/**
 * 新卒面接共有カレンダー登録用のヘッダー行のインデックスを取得する
 * 
 * @param  {Array.<Array.<string|date>>} values - シートの値、2次元配列
 * @return {Object.<number>}
 * 
 */
function assignHeaderKeys_(values, rowIndex){

  console.info(`assignHeaderKeys_()を実行中`);
  console.info(`02_calenderに記載`);

  const headers = values[rowIndex];
  const keys    = ['eventId', 'title', 'date', 'startTime', 'endTime', 'attendees', 'description', 'status'];
  const array   = ['イベントID', 'イベント名', 'イベント予定日', '開始時刻', '終了時刻', '出席者', 'イベント詳細', '登録ステータス' ];
  const column  = buildObjectFromPairs(headers, keys, array);
 
  return column
}



/**
 * Google Meet付きで予定を登録し、Meet URLを返す
 * 
 * @param  {Object.<string>} object - 予定、日時、詳細などの情報
 * @return {string} 
 * 
 */
function createEventWithMeetUrl_(object) {

  console.info(`createEventWithMeetUrl_()を実行中`);
  console.info(`02_calenderに記載`);
  
  const eventParam = {
    conferenceData: {
      createRequest: {
        conferenceSolutionKey: {
          type: "hangoutsMeet"
        },
        requestId: "123"
      }
    },
    summary: object.title,//カレンダータイトル
    description: object.description,
    start:     {dateTime: object.startTime},
    end:       {dateTime: object.endTime},
    attendees: object.attendees,
  };
 
  //CalendarAPIに対し、Meet会議付き予定を追加
  const event = Calendar.Events.insert(eventParam, object.calId, {conferenceDataVersion: 1});
  console.log('登録成功');
  console.log(`イベントID：${event.id}`);
 
  return event.id
 
}



/**
 * 'attendees': [{'email': 'lpage@example.com'},{'email': 'sbrin@example.com'}],
 * registerEventsWithMeetUrl_()内で、上記のように指定すると出席者を登録する事が出来る
 *
 * @param  {string} string - 'aiueo@sample.com, abc@sample.com'
 * @return {Array.<Object.<string>}
 * 
 */
function generateAttendees_(string){

  console.info(`generateAttendees_()を実行中`);
  console.info(`02_calenderに記載`);

  const newArray = string.split(',').map(email => ({'email': email}));
  console.log(newArray);
  return newArray
}



/**
 * 予定を削除する、削除前にはアラートが表示される
 * 
 * @param {string} query - 削除したい予定のタイトル　省略可。省略した場合はプロンプトに入力する
 * @param {string} calId - カレンダーID  省略可。　省略した場合は、自分のアカウントで処理が実施される。
 * 
 */
function deleteEvents(query, calId) {

  console.info(`deleteEvents()を実行中`);
  console.info(`02_calenderに記載`);

  if(!query) query = showPrompt('削除したい予定名を入力してください', '（例）：テスト');

  // 削除対象の予定情報を取得し、eventIdArray を作成
  const events = getEventsByTitle_(query, calId);

  // ダイアログを表示してユーザーに確認を求める
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `${query}　を含む予定が${events.array.length}件あります。
    \n削除してもよろしいですか？\n\n
    ${events.string}`, ui.ButtonSet.YES_NO);

  switch (response) {
    case ui.Button.YES:
      console.log('“はい” のボタンが押されました。');
      events.array.forEach(eventId => CalendarApp.getEventById(eventId).deleteEvent());
      ui.alert('予定の削除が完了しました。');
      break;

    case ui.Button.NO:
      console.log('“いいえ” のボタンが押されました。');
      ui.alert('処理が中断されました。');
      break;

    default:
      console.log('処理が中断されました。');
      return;
  }
}



/**
 * eventIdと削除予定のイベントを含んだオブジェクトを返す
 * 
 * @param  {string} query - 削除したいカレンダーの予定名
 * @param  {string} calId - カレンダーID
 * @return {Object.<Array.<string> | string>}
 * 
 */
function getEventsByTitle_(query, calId){

  console.info(`()を実行中`);
  console.info(`02_calenderに記載`);

  const cal       = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
  const startDate = new Date();
  const endDate   = new Date();
  endDate.setMonth(endDate.getMonth() + 1);

  let string = '';

  // filterメソッドを使用し、undefinedが配列に追加されることを防ぐ
  const eventIdArray = cal.getEvents(startDate, endDate)
    .filter(event => event.getTitle().includes(query))
    .map((event, index) => {
      string += `${index + 1}. ${event.getTitle()}  ${formatDate(event.getStartTime(), 'yyyy/MM/dd (E) HH:mm')}\n`;
      return event.getId();
    });

  const eventsInfo = {
    array:  eventIdArray,
    string: string
  };

  console.log(eventsInfo);
  return eventsInfo
  
}



/**
 * 予定を編集するスクリプトを起動する画面を表示する
 * FIXME: ライブラリ経由だとHTMLのformタグの内容が受け取れない 値がnullになる模様
 * 
 */
function showEditEventsLauncher() {
  const html = HtmlService.createHtmlOutputFromFile('choice');
  SpreadsheetApp.getUi().showModalDialog(html, 'どのように編集したいかを選択してください');
  //choice.htmlでGoogle.script.runが動くはず
  //返り値などは必要ない
}



/**
 * choice.htmlで選択した内容を受け取って、編集内容を分岐させる
 * 
 * @param  {Object.<string>} object - 選択肢の内容　以下4つ
 * (例)予定名を編集する, 詳細欄を編集する, 日時を編集する, 出席者を追加する
 * 
 */
function editEvents(object, rowIndex){

  console.info(`editEvents()を実行中`);
  console.info(`02_calenderに記載`);

  //FIXME: google.script.run.editEvents()の引数に2個指定することが出来ない模様
  //引数にrowIndexがundefinedであれば0を代入する
  rowIndex = rowIndex ? rowIndex : 0;

  // 引数 'object' の型によって条件分岐して文字列に変換する
  const eventItem = typeof object === 'object' ? String(Object.values(object)) : String(object);
  const calId     = showPrompt('カレンダーIDを入力してください', '空白の場合は自分のカレンダーを処理対象とします');
  const cal       = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();

  console.log(`eventItem: ${eventItem} typeOf ${typeof eventItem}`);
  console.log(`カレンダーID：${cal.getId()}`);
  console.log(`処理対象のアカウント：${cal.getName()}`);
  
  //シート上を走査して、編集対象のイベントIDを取得する
  const sheet  = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const column = assignHeaderKeys_(values, rowIndex);

  let count = 0;

  values.forEach((array, index) => {
    if (array[column.status] !== '編集対象') return;

    const updateInfo = {
      event:       cal.getEventById(array[column.eventId]),
      title:       array[column.title],
      date:        array[column.date],
      start:       array[column.startTime],
      end:         array[column.endTime],
      description: array[column.description],
      guests:      array[column.attendees].split(',')
    }
      
    switch(eventItem){
      case '予定名を編集する':
        event.setTitle(updateInfo.title);
        break;

      case '詳細欄を編集する':
        event.setDescription(updateInfo.description);
        break;

      case '日時を編集する':
        updateEventDateTime(updateInfo);  
        break;

      case '出席者を追加する':
        updateInfo.guests.forEach(guest => event.addGuest(guest));
        break;

      default:
        console.log('該当しませんでした');
    }

    const row = index + 1;
    sheet.getRange(row, column.status + 1).setValue('編集済');
    console.log(`${row} 行目  予定名：${updateInfo.title}`);

    count += 1;
  });
  SpreadsheetApp.getUi().alert(`${count}件の予定を変更しました`);
}



/**
 * イベントIDから予定名、日時などの詳細を取得する
 * 
 * @param  {string} eventId - イベントID、UI操作だと確認不可
 * @param  {string} calId - 予定の登録先、通常はメールアドレスなど
 * @return {Object.<string>}　予定の詳細
 * 
 */
function getInfoByEventId(eventId, calId){

  console.info(`getInfoByEventId()を実行中`);
  console.info(`02_calenderに記載`);

  const cal   = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
  const event = cal.getEventById(eventId);
  const info  = {
    eventTitle:  event.getTitle(),
    targetDate:  formatDate(event.getStartTime(), 'yyyy/MM/dd'),
    targetDay:   convertDay(event.getStartTime().getDay()),
    startTime:   formatDate(event.getStartTime(), 'HH:mm'),
    endTime:     formatDate(event.getEndTime(), 'HH:mm'),
    eventDetail: event.getDescription()
  }

  console.log(`処理対象のアカウント：${cal.getName()}`);
  console.log(info);

  return info;
}



/**
 * 予定の日時を更新する
 * 
 * @param  {Object.<date>} info - 開始時刻、終了時刻を含むオブジェクト
 * 
 */
function updateEventDateTime(info){

  console.info(`updateEventDateTime()を実行中`);
  console.info(`02_calenderに記載`);

  const startTime = new Date(info.date);
  startTime.setHours(info.start.getHours());
  startTime.setMinutes(info.start.getMinutes());

  const endTime = new Date(startTime);
  endTime.setHours(info.end.getHours());
  endTime.setMinutes(info.end.getMinutes());

  info.event.setTime(startTime, endTime);
}



/**
 * 祝日かどうかを判定する
 * 
 * @param  {string|date} date - 文字列かデイトオブジェクト
 * @return {boolean}
 * 
 */
function isHoliday(date){

  console.info(`isHoliday()を実行中`);
  console.info(`02_calenderに記載`);

  date = typeof date === 'string' ? date = new Date(date) : date;
  const targetDate = formatDate(date, 'yyyy/MM/dd');
  const values     = getHolidays_();
  const holidays   = generateArray(values, 1); 
  const foundIndex = holidays.findIndex(holiday => holiday === targetDate);

  (foundIndex !== -1) ? console.log(`${targetDate}は、${values[foundIndex][0]}で祝日です`)
  : console.log(`${targetDate}は、祝日ではありません`);
  return foundIndex !== -1;

}



/**
 * 祝日カレンダーから祝日名、日付、曜日を取得する
 * 
 */
function getHolidays_(){

  console.info(`getHolidays_()を実行中`);
  console.info(`02_calenderに記載`);

  const cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  const startDate = new Date();
  const endDate   = new Date();
  endDate.setFullYear(endDate.getFullYear()+1);

  const dayOfWeek = '日月火水木金土';
  const daysArray = dayOfWeek.split('');
  const day = (number) => {
    return daysArray[number];
  }

  const events = cal.getEvents(startDate, endDate).map(event => ({
    title: event.getTitle(),
    date:  Utilities.formatDate(event.getStartTime(), 'JST', 'yyyy/MM/dd'),
    day:   day(event.getStartTime().getDay())
    })
  );

  const keys   = Object.keys(events[0]);
  const values = events.map(event => keys.map(key => event[key]));
  values.unshift(['祝日名', '日付', '曜日']);

  console.log(values);
  return values
}