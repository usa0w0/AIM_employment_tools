// formを開く関数
function penform(){
  // スクリプトプロパティからformのIDを取得
  const form_ID = PropertiesService.getScriptProperties().getProperty("form_ID");
  // 応募フォームを取得
  var form = FormApp.openById(form_ID);

  // 回答を受け付ける
  form.setAcceptingResponses(true);
}

//フォームを閉じる関数
function closeform(){
  // スクリプトプロパティからformのIDを取得
  const form_ID = PropertiesService.getScriptProperties().getProperty("form_ID");
  // 応募フォームを取得
  var form = FormApp.openById(form_ID);

  // 締切メッセージのために、年度と学期を取得
  // 実行時の日付から、年度と学期を取得
  const now_date = new Date();
  var month = now_date.getMonth();
  if (8 < month){
    year = now_date.getFullYear()+1;
    semester = "前期";
  } else if (month < 4){
    year = now_date.getFullYear();
    semester = "前期";
  } else {
    year = now_date.getFullYear();
    semester = "後期";
  }

  // 回答を締め切る
  var custom_close_message = year + "年度" + semester + "のAIM新規学生スタッフ募集は終了しました。応募者の方は、後ほどAoyama-mailにご連絡いたしますのでお待ちください。別途、質問等ある方は、sagamipro-contact@aim.aoyama.ac.jpまでご連絡ください。"
  form.setCustomClosedFormMessage(custom_close_message);
  form.setAcceptingResponses(false);
}