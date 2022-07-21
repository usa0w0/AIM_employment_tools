// メール送信機能をメニュー追加する関数
function form_ss_addMenu(){
  // getActiveでメニュー追加対象のSSを取得
  // スクリプトプロパティからformのIDを取得
  const form_ss_ID = PropertiesService.getScriptProperties().getProperty("Form_SS_ID");
  // 応募フォームを取得
  var form_ss = SpreadsheetApp.openById(form_ss_ID);

  // 集計SSへ追加する機能メニュー
  form_ss_menu = [
    {name: "面接日時取得", functionName: "getInterviewDate"},
    {name: "メール自動送信", functionName: "SendAsMail"}
  ];

  form_ss.addMenu("採用ツール", form_ss_menu);
}

// 関数実行のテスト用
function sendMail() {
  //メッセージボックスを出す
  var result = Browser.msgBox("<テストメッセージ> 追加された機能が実行されます。", Browser.Buttons.OK);
}