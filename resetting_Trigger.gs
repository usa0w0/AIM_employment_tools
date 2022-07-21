function resetTriger() {
  // トリガーを全削除
  var allTriggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < allTriggers.length; i++) {
    console.log(allTriggers[i])
    ScriptApp.deleteTrigger(allTriggers[i]);
  }

  var form_ss_ID = property.getProperty("Form_SS_ID");
  if (form_ss_ID == null){
    Browser.msgBox("集計SSが生成されていません。\n（スクリプトプロパティに存在しません。）", Browser.Buttons.OK);
  } else {
    // form_ssに機能メニューの追加
    ScriptApp.newTrigger("form_ss_addMenu")
    .forSpreadsheet(form_ss)
    .onOpen().create();
  }

  var form_ID = property.getProperty("Form_ID");
  if (form_ID == null){
    Browser.msgBox("応募フォームが生成されていません。\n（スクリプトプロパティに存在しません。）", Browser.Buttons.OK);
  } else {
    changeManageSche()
  }
}

function changeManageSche(){
  // スクリプトプロパティの募集開始・終了日を更新
  setNewProperty();

  // 自動で回答の受付・締切を制御するトリガーをセット
  // 指定の日付と時刻でトリガーをセットする
  ScriptApp.newTrigger("openform")
  .timeBased()
  .at(open_date)
  .create();
  
  ScriptApp.newTrigger("closeform")
  .timeBased()
  .at(close_date)
  .create();
}