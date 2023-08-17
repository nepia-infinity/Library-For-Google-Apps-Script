/**
 * dateオブジェクトを、yyyy/MM/ddなどの指定した文字列で返す。
 * E - 曜日の指定があった際に　wed → 水　のように変換する
 * 
 * @param  {date} date - dateオブジェクト
 * @param  {sting} format - 'yyyy/MM/dd', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd (E)'
 * @return {string} （例）2022/04/06
 */
function formatDate(date, format){

  console.info(`formatDate()を実行中`);
  console.info(`02_calenderに記載`);

  const formatDate = Utilities.formatDate(date, 'JST', format);

  if(formatDate.match(/[a-zA-Z]/)!== null){

    // 2022/04/06 (wed) → 2022/04/06 (水)　のように変換
    const day    = convertDay(date.getDay());
    const marged = formatDate.replace(/[a-zA-Z]{3}/, `${day}`);

    console.log(`変換前の表記：　${formatDate}`);
    console.log(`変換後の表記：　${marged}`);
    return marged

  }else{
    //曜日の指定がない場合
    console.log(`日付　：　${formatDate}`);
    return formatDate
  }
}


/**
 * 
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

  let day, stringDate;

  let newValues = [];

  if(targetDate < today){
    // targetDateが、今日の日付より過去の場合
    for(let d = targetDate; d < today; d.setDate(d.getDate()+1)) {
      day        = convertDay(d.getDay());
      stringDate = formatDate(d, 'yyyy/MM/dd');
      newValues.push([stringDate, day]);

    }//for
  }else if(today < targetDate){
    // targetDateが、今日の日付より未来の場合
    for(let d = today; d < targetDate; d.setDate(d.getDate()+1)){
      day        = convertDay(d.getDay());
      stringDate = formatDate(d, 'yyyy/MM/dd');
      newValues.push([stringDate, day]);
    }
  }
  console.log(newValues);
  return newValues
}//end



/**
 * 英語表記の曜日、sat, wedなどを、'日月火水木金土'のいずれかに変換する
 * 
 * @param  {number} tempNumber - date.getDay()
 * @return {string} 
 * 
 */
function convertDay(tempNumber){

  console.info(`convertDay()を実行中`);
  console.info(`02_calenderに記載`);

  // 文字列を配列化
  const dayOfWeek = '日月火水木金土';
  const daysArray = dayOfWeek.split('');
  const day       = daysArray[tempNumber];

  console.log(daysArray);
  console.log(`daysArray[${tempNumber}]　${day}曜日`);

  return day

}



/**
 * 引数にした日時からnカ月後の予定を2次元配列で取得する
 * FIXME:全ての予定を取得しているため、予定がたくさん登録されている場合は　を使用してください
 * 
 * @param  {string} string - 'yyyy/MM/dd'形式で日付を指定　（例）'2023/04/28'
 * @param  {number} offset - nカ月後にあたる　数値で指定
 * @param  {string} calId - カレンダーID
 * @return {Array.<Array.<string>>} 
 * 
 */
function getCalAllEvents(string, offset, calId) {

  console.info(`convertDay()を実行中`);
  console.info(`02_calenderに記載`);

  let cal;

  if(calId){
    // 特定のカレンダーを取得する
    cal = CalendarApp.getCalendarById(calId);

  }else{
    // 3つ目の引数の指定がない場合、自分のカレンダーを取得する
    console.warn(`3番目の引数の指定がないため、あなた自身のカレンダーの予定を取得します。`);
    cal = CalendarApp.getDefaultCalendar();

  }
        
  const startTime = new Date(string);
  const endTime   = new Date();
  endTime.setMonth(endTime.getMonth() + offset);

  //指定のカレンダーから予定を取得
  const events = cal.getEvents(startTime, endTime).map(event => ({
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

  console.log(`オブジェクトを2次元配列に変換`);
  console.log(values);
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
function guestList_(guests, creators){

  console.info(`guestList_()を実行中`);
  console.info(`02_calenderに記載`);
  
  let array = [];

  if(creators){
    array.push(creators[0]); //主催者を追加

  }
  guests.map(guest => array.push(guest.getEmail()));

  //配列を文字列化する
  const guestsList = array.join();
  return guestsList

}



/**
 * シートの値を元にカレンダーに登録する
 * FIXME: Google Calender APIを有効にしておかないとエラーが生じてしまう。
 * FIXME: 参加者欄が空白だとエラーが生じる
 * 
 * @param {string} url - スプレッドシートのURL
 * 
 */
function createEventsFromSheetValues(url) {

  console.info(`createEventsFromSheetValues()を実行中`);
  console.info(`02_calenderに記載`);

  const sheet  = getSheetByUrl(url);
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  
  const column = {
    eventId:     header.indexOf('イベントID'),
    title:       header.indexOf('イベント名'),
    date:        header.indexOf('イベント予定日'),
    startTime:   header.indexOf('開始時刻'),
    endTime:     header.indexOf('終了時刻'),
    attendees:   header.indexOf('出席者'),
    description: header.indexOf('イベント詳細'),
    status:      header.indexOf('登録ステータス')
  }

  let count = 0;
  
  for(let i = 0; i < values.length; i++){
    //二重登録防止のため、イベントIDが空白かつ、登録済みが付いていない予定のみを登録する。
    if(!values[i][column.eventId] && !values[i][column.status]){
      
      //開始時刻
      const startTime = new Date(values[i][column.date]);
      startTime.setHours(values[i][column.startTime].getHours());
      startTime.setMinutes(values[i][column.startTime].getMinutes());
      
      //終了時刻
      const endTime = new Date(startTime);
      endTime.setHours(values[i][column.endTime].getHours());
      endTime.setMinutes(values[i][column.endTime].getMinutes());
      
      const row = i + 1;
      console.log(`処理中：${row}　行目`);
      
      const object = {
        title:       values[i][column.title],
        startTime:   Utilities.formatDate(startTime, 'JST', "yyyy-MM-dd'T'HH:mm:ss.000+09:00"),
        endTime:     Utilities.formatDate(endTime, 'JST', "yyyy-MM-dd'T'HH:mm:ss.000+09:00"),
        description: values[i][column.description],
        attendees:   generateAttendees_(values[i][column.attendees]),
      }

      const eventId = registerEventsWithMeetUrl_(object);

      sheet.getRange(row, column.eventId + 1).setValue(eventId);
      sheet.getRange(row, column.status  + 1).setValue('登録済');

      count += 1;

    }
  }
  console.log(`登録数：　${count}　件`);
  SpreadsheetApp.getUi().alert(`${count}件の登録が完了しました。カレンダーをご確認ください。`);
}




/**
 * Meet URL付きのカレンダーを登録する
 * https://auto-worker.com/blog/?p=6252
 * 
 * @param  {Object.<number>}　{id:0, title:1}のようなオブジェクトで設定
 * @param  {string}　calId - 登録するアカウントID 大抵の場合はメールアドレス
 * @return {string} 新規登録された予定のイベントID
 * 
 */
function registerEventsWithMeetUrl_(object, calId) {

  if(!calId) calId = Session.getActiveUser().getEmail();
  console.log(`登録用アカウント：　${calId}`);

  //GoogleカレンダーでMeet会議が設定されるイベント登録パラメータを設定
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
    start: {
      dateTime: object.startTime
    },
    end: {
      dateTime: object.endTime
    },
    attendees: object.attendees,
  };

  //CalendarAPIに対し、Meet会議付き予定を追加
  const event = Calendar.Events.insert(eventParam, calId, {conferenceDataVersion: 1});
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
 */

function generateAttendees_(string){
  const newArray = string.split(',').map(email => ({'email': email}));
  console.log(newArray);
  return newArray
}



/**
 * 
 * 予定を削除する、削除前にはアラートが表示される
 * 
 * @param {string} query - 削除したい予定のタイトル　省略可。省略した場合はプロンプトに入力する
 * @param {string} calId - カレンダーID  省略可。　省略した場合は、自分のアカウントで処理が実施される。
 * 
 */
function deleteEvents(query, calId){

  if(!calId) calId = Session.getActiveUser().getEmail();
  console.log(`登録用アカウント：　${calId}`);

  if(!query) query = showPrompt('削除したい予定名を入力してください', '（例）：テスト');

  const cal       = CalendarApp.getCalendarById(calId);
  const startDate = new Date();
  const endDate   = new Date();
  endDate.setMonth(endDate.getMonth() + 1);

  let string       = '';
  let count        = 0;
  let eventIdArray = [];

  cal.getEvents(startDate, endDate).map(event => {
    const info = {
      eventId:    event.getId(),
      title  :    event.getTitle(),
      targetDate: formatDate(event.getStartTime(), 'yyyy/MM/dd HH:mm'),
      targetDay: convertDay(event.getStartTime().getDay())
    }

    // queryを含んでいなかったら処理をスキップ
    if(!info.title.includes(query)) return;
    count  += 1;
    string += `${count}. ${info.title} (${info.targetDay}) ${info.targetDate}\n`; 
    eventIdArray.push(info.eventId);
  });

  console.log(string);
  console.log(`該当する予定が${count}件あります。`);

  const ui       = SpreadsheetApp.getUi();
  const response = ui.alert(`該当する予定が${count}件あります。\n
    削除してもよろしいですか？\n\n
    ${string}`, ui.ButtonSet.YES_NO
  );

  switch (response){
    case ui.Button.YES:
      console.log('“はい” のボタンが押されました。');
      eventIdArray.map(eventId => CalendarApp.getEventById(eventId).deleteEvent());
      ui.alert('予定の削除が完了しました。');
      break;

    case ui.Button.NO:
      console.log('“いいえ” のボタンが押されました。');
      ui.alert('処理が中断されました。');
      break;

    default:
      console.log('処理が中断されました。');
      return
  }
}



/**
 * 
 * 予定を編集するスクリプトを起動する画面を表示する
 * 
 */
function showEditEventsLauncher() {
  const html = HtmlService.createHtmlOutputFromFile('choice');
  SpreadsheetApp.getUi().showModalDialog(html, 'どのように編集したいかを選択してください');
}


/**
 * choice.htmlで選択した内容を受け取って、編集内容を分岐させる
 * @param  {Object.<string>} object - 選択肢の内容 
 *
 * 
 */
function editEvents(object){
  const argument = String(Object.values(object));
  console.log(`argument: ${argument} typeOf ${typeof argument}`);

  const calId = showPrompt('カレンダーIDを入力してください', '空白の場合は自分のカレンダーを処理対象とします');
  console.log(`カレンダーID：　${calId}`);

  let cal;

  if(calId){
    cal = CalendarApp.getCalendarById(calId);
    console.log(`処理対象のアカウント：　${cal.getName()}`);

  }else {
    cal = CalendarApp.getDefaultCalendar();
    console.log(`処理対象のアカウント：　${cal.getName()}`);

  }
  
  //シート上を走査して、編集対象のイベントIDを取得する
  const sheet  = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const column = {
    eventId:     header.indexOf('イベントID'),
    title:       header.indexOf('イベント名'),
    date:        header.indexOf('イベント予定日'),
    startTime:   header.indexOf('開始時刻'),
    endTime:     header.indexOf('終了時刻'),
    attendees:   header.indexOf('出席者'),
    description: header.indexOf('イベント詳細'),
    status:      header.indexOf('登録ステータス')
  }

  let count = 0;

  for(let i = 0; i < values.length; i++){
    if(values[i][column.status] === '編集対象'){

      const event = cal.getEventById(values[i][column.eventId]);
      console.log(`予定名：${event.getTitle()}`);

      switch(argument){
        case '予定名を編集する':
          event.setTitle(values[i][column.title]);
          break;
        case '詳細欄を編集する':
          event.setDescription(values[i][column.description]);
          break;
        case '日時を編集する':
          const startTime = new Date(values[i][column.date]);
          startTime.setHours(values[i][column.startTime].getHours());
          startTime.setMinutes(values[i][column.startTime].getMinutes());

          const endTime = new Date(startTime);
          endTime.setHours(values[i][column.endTime].getHours());
          endTime.setMinutes(values[i][column.endTime].getMinutes());

          event.setTime(startTime, endTime);
          break;
        case '出席者を追加する':
          const guests = values[i][column.attendees].split(',');
          console.log(guests);

          guests.map(guest => event.addGuest(guest));
          break;
        default:
          console.log('該当しませんでした');
      }

      const row = i + 1;
      console.log(`処理対象行：　${row}`);

      sheet.getRange(row, column.status + 1).setValue('編集済');
      count += 1;

    }else {
      continue;
    }
  }
  SpreadsheetApp.getUi().alert(`${count}件の予定を変更しました`);
}


/**
 * @param  {string} eventId - イベントID、UI操作だと確認不可
 * @param  {string} calId - 予定の登録先、通常はメールアドレスなど
 * @return {Object.<string>}　予定の詳細
 * 
 */
function getInfoByEventId(eventId, calId){
  let cal;
  if(calId){
    cal = CalendarApp.getCalendarById(calId);

  }else{
    cal = CalendarApp.getDefaultCalendar();

  }

  console.info(`getInfoByEventId()を実行中`);
  console.info(`02_calenderに記載`);

  const event = cal.getEventById(eventId);
  const info  = {
    eventTitle:  event.getTitle(),
    targetDate:  formatDate(event.getStartTime(), 'yyyy/MM/dd'),
    targetDay:   convertDay(event.getStartTime().getDay()),
    startTime:   formatDate(event.getStartTime(), 'HH:mm'),
    endTime:     formatDate(event.getEndTime(), 'HH:mm'),
    eventDetail: event.getDescription()
  }

  console.log(info);
  return info;
}