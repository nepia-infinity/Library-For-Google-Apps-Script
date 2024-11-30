/**
 * 検索キーワードに一致する内容のメールを2次元配列で取得する
 * ヘッダー行　['受信日', '送信者', '件名', '本文', 'HTML']　
 * 
 * @param  {string} query - 検索したい語句
 * @param  {number} quantity - 引数に100とした場合、100件のスレッドを検索する　
 * @return {Array.<Array.<string>>}
 */
function getGmailThread(query, quantity){

  console.info('getGmailThread()を実行中');
  console.info('05_Gmailに記載');
 
  const threads = GmailApp.search(query, 0, quantity);
  let values    = [[ '受信日', '送信者', '件名', '本文', 'HTML']];

  console.log(`検索キーワード：　${query} を　${quantity}件、検索します`);

  //スレッドを取得する
  for (const thread of threads){
    const messages = thread.getMessages();

    //一つ一つのスレッドにある各メッセージを取得する。
    for(const message of messages){
      const body    = message.getPlainBody();
      const subject = message.getSubject();

      // 件名もしくは内容に検索queryを含む場合のみ処理を進める
      if(body.includes(query) || subject.includes(query)){
        const info = {
          date:    formatDate(message.getDate(), 'yyyy/MM/dd HH:mm:ss'),
          from:    message.getFrom(),
          subject: subject,
          text:    body,
          html:    message.getBody(),
        }
        values.push(Object.values(info));
      }
    }//for
  }//for

  console.log(values);
  return values

}//end