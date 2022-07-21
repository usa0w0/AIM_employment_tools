function createTools() {
  // スクリプトプロパティをクリア
  property.deleteAllProperties();
  // プロパティにパラメータをセット
  setNewProperty()

  // フォームと集計SSを生成して、回答を連携させる
  createSS()
  Logger.log("SS生成完了")

  createForm()
  Logger.log("フォーム生成完了")

  linkFormToSS()
  Logger.log("回答連携完了")
}