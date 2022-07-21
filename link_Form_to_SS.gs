function linkFormToSS() {
  // 集計SSと連携
  // スクリプトプロパティからformのIDを取得
  const form_ID = PropertiesService.getScriptProperties().getProperty("form_ID");
  // 応募フォームを取得
  var form = FormApp.openById(form_ID);
  // スクリプトプロパティからform_ssのIDを取得
  const form_ss_ID = property.getProperty("Form_SS_ID")
  form.setDestination(FormApp.DestinationType.SPREADSHEET, form_ss_ID);
}